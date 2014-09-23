#Region "Namespaces"
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
#End Region

<Assembly: CommandClass(GetType(WJL1BlockTools.Commands))> 
<Assembly: ExtensionApplication(GetType(WJL1BlockTools.Initialisation))> 

Namespace WJL1BlockTools

    Public Class Initialisation
        Implements Autodesk.AutoCAD.Runtime.IExtensionApplication

        Dim BlockTools As New BlockTools
        Public Sub Initialize() Implements IExtensionApplication.Initialize
            'Dim myDWG As Document
            'Dim myDB As Database = HostApplicationServices.WorkingDatabase()
            'myDWG = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim myEd As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            myEd.WriteMessage(vbLf & "Initialising BlockTools commands...")
            myEd.WriteMessage(vbLf & "   'BlockRealign' realigns blocks")
        End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate

        End Sub
    End Class

    Public Class Commands
        Dim LayerToRestore As String = ""
        Dim BlockTools As New BlockTools

        <CommandMethod("BlockRealign")> _
        Public Sub BlockRealign_Method()
            BlockTools.BlockRealign()
        End Sub

        <CommandMethod("SuperExplode")> _
        Public Sub SuperExplode_Method()
            BlockTools.SuperExplode()
        End Sub


    End Class

    Public Class BlockTools
        Enum RealignOpts As Integer 'Enum for layer state options
            None = 0
            Name = 1
            Origin = 2
            Rotation = 4
            Scale = 8
        End Enum


        Sub Settings()
            Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
            Dim myDB As Database = myDWG.Database
            Dim myEd As Editor = myDWG.Editor

            'Dim RealignOpts As BlockTools.RealignOpts = WJL1BlockTools.BlockTools.RealignOpts.Name Or _
            '                                                                WJL1BlockTools.BlockTools.RealignOpts.Origin Or _
            '                                                                WJL1BlockTools.BlockTools.RealignOpts.Rotation Or _
            '                                                                WJL1BlockTools.BlockTools.RealignOpts.Scale
            Dim RealignOpts As BlockTools.RealignOpts = My.Settings.BlockRealignMode

            myEd.WriteMessage(vbLf & "BlockRealign modes:" & _
                              vbLf & "Name = 1" & _
                              vbLf & "Origin = 2" & _
                              vbLf & "Rotation = 4" & _
                              vbLf & "Scale = 8")

            ''Get the realign mode from the user (as bitwise integer)
            Dim GetIntegerMessage As String = vbLf & "Enter new mode for BlockRealign"
            Dim GetIntegerOpts As New PromptIntegerOptions(GetIntegerMessage)
            GetIntegerOpts.DefaultValue = RealignOpts
            GetIntegerOpts.UseDefaultValue = True
            GetIntegerOpts.LowerLimit = 1
            GetIntegerOpts.UpperLimit = 15

            Dim GetIntegerResult As PromptIntegerResult = myEd.GetInteger(GetIntegerOpts)

            Select Case GetIntegerResult.Status
                Case PromptStatus.OK
                    My.Settings.BlockRealignMode = GetIntegerResult.Value
                    My.Settings.Save()
            End Select

        End Sub

        Sub BlockRealign()

            Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
            Dim myDB As Database = myDWG.Database
            Dim myEd As Editor = myDWG.Editor

            Dim curUCSMatrix As Matrix3d = myDWG.Editor.CurrentUserCoordinateSystem
            Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d

KEYWORD_ENTERED:
            ''Set the GetEntity options
            Dim GetEntOpts As New PromptEntityOptions("")

            ''Add the Keywords & Message
            GetEntOpts.Keywords.Add("Settings")
            GetEntOpts.Message = vbLf & "Select Block to realign or [Settings]: "

            GetEntOpts.AppendKeywordsToMessage = False

            ''Filter so only block references get selected
            GetEntOpts.SetRejectMessage(vbLf & "Selected entity is not a Block.")
            GetEntOpts.AddAllowedClass(GetType(BlockReference), False)

            Dim GetEntResult As PromptEntityResult = myEd.GetEntity(GetEntOpts)

            Select Case GetEntResult.Status
                Case PromptStatus.Keyword
                    Select Case GetEntResult.StringResult
                        Case "Settings"
                            Settings()
                            GoTo KEYWORD_ENTERED
                        Case Else
                            MsgBox("Unhandled keyword: " & GetEntResult.StringResult)
                    End Select
                Case PromptStatus.OK
                    Dim myTrans As Transaction = myDWG.TransactionManager.StartTransaction
                    Try
                        Dim SelectedObject As Object = myTrans.GetObject(GetEntResult.ObjectId, OpenMode.ForRead)
                        Dim SelectedEntity As Entity = CType(SelectedObject, Entity)

                        Dim BlockRef As BlockReference
                        BlockRef = SelectedEntity



                        Dim BlockDefinition As BlockTableRecord = myTrans.GetObject(BlockRef.BlockTableRecord, OpenMode.ForRead)
                        If BlockDefinition.IsFromExternalReference Then
                            myEd.WriteMessage(vbLf & "Selected entity is not a Block.")
                        Else
                            If BlockDefinition.IsDynamicBlock Then
                                myEd.WriteMessage(vbLf & "Selected entity is a Dynamic block.")
                            Else


                                ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                                ''Change the block name:
                                If (My.Settings.BlockRealignMode Or WJL1BlockTools.BlockTools.RealignOpts.Name) = My.Settings.BlockRealignMode Then
RETRY_NAME_CHANGE:
                                    Dim GetStringMessage As String = vbLf & "Enter new name for block '" & BlockRef.Name & "' "
                                    Dim GetStringOpts As New PromptStringOptions(GetStringMessage)
                                    GetStringOpts.DefaultValue = BlockRef.Name
                                    GetStringOpts.UseDefaultValue = True
                                    GetStringOpts.AllowSpaces = True

                                    Dim GetStringResult As PromptResult = myEd.GetString(GetStringOpts)

                                    Select Case GetStringResult.Status
                                        Case PromptStatus.OK
                                            If GetStringResult.StringResult <> "" Then
                                                Try
                                                    BlockDefinition.UpgradeOpen()
                                                    BlockDefinition.Name = GetStringResult.StringResult
                                                Catch ex As Exception
                                                    myEd.WriteMessage(vbLf & "Unable to change name of block '" & BlockRef.Name & "' to '" & GetStringResult.StringResult & "'.")
                                                    GoTo RETRY_NAME_CHANGE
                                                End Try

                                            End If
                                        Case Else
                                            myTrans.Abort()
                                            GoTo CANCEL
                                    End Select
                                End If
                                ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

                                ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                ''Change the block origin:
                                If (My.Settings.BlockRealignMode Or WJL1BlockTools.BlockTools.RealignOpts.Origin) = My.Settings.BlockRealignMode Then
                                    ''Get the current insertion point of the selected block
                                    Dim BlockRefInPoint As Point3d = BlockRef.Position

                                    ''Get the new insertion point from the user
                                    Dim GetPointMessage As String = vbLf & "Enter new insertion point for block '" & BlockRef.Name & "' <" & Math.Round(BlockRefInPoint.X, 6) & "," & Math.Round(BlockRefInPoint.Y, 6) & "," & Math.Round(BlockRefInPoint.Y, 6) & "> "
                                    Dim GetPointOpts As New PromptPointOptions(GetPointMessage)
                                    GetPointOpts.BasePoint = BlockRefInPoint
                                    GetPointOpts.UseBasePoint = True
                                    GetPointOpts.UseDashedLine = True
                                    GetPointOpts.AllowArbitraryInput = False
                                    GetPointOpts.AllowNone = True
                                    Dim GetPointResult As PromptPointResult = myEd.GetPoint(GetPointOpts)

                                    Select Case GetPointResult.Status
                                        Case PromptStatus.None
                                            'Use current point, don't do anything
                                        Case PromptStatus.OK
                                            ''If the new insertion point differs from the old insertion point (waste of time otherwise)
                                            If GetPointResult.Value <> BlockRefInPoint Then
                                                ''Calculate the InpointVector
                                                Dim InpointVector As Vector3d = BlockRefInPoint - GetPointResult.Value

                                                ''Rotate the InpointVector to account for the selected block's rotation
                                                InpointVector = InpointVector.TransformBy(Matrix3d.Rotation((BlockRef.Rotation * -1), curUCS.Zaxis, GetPointResult.Value))
                                                ''Rotate the InpointVector to account for the selected block's scale
                                                InpointVector = InpointVector.TransformBy(Matrix3d.Scaling((1 / BlockRef.ScaleFactors.X), BlockRefInPoint))

                                                ''Redefine the block, moving everything to match the new insertion point
                                                BlockDefinition.UpgradeOpen()
                                                For Each myObjID As ObjectId In BlockDefinition
                                                    ''Move each object in the block definition
                                                    Dim myEnt As Entity = myObjID.GetObject(OpenMode.ForWrite)
                                                    myEnt.TransformBy(Matrix3d.Displacement(InpointVector))
                                                Next

                                                ''Get ModelSpace 
                                                Dim myBT As BlockTable = myDB.BlockTableId.GetObject(OpenMode.ForRead)
                                                Dim myModelSpace As BlockTableRecord = myBT(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForRead)

                                                'For Each Object in ModelSpace
                                                For Each myObjID As ObjectId In myModelSpace
                                                    Dim myEnt As Entity = myObjID.GetObject(OpenMode.ForRead)

                                                    ''Check if the entity is a block refernce
                                                    If TypeOf myEnt Is BlockReference Then
                                                        Dim myBlockRef As BlockReference = myEnt

                                                        ''Check that it's the right block
                                                        If myBlockRef.Name = BlockDefinition.Name Then
                                                            ''Calculate the vector to move the block reference, accounting for scale & rotation respectively
                                                            Dim myInPointVector As Vector3d
                                                            myInPointVector = InpointVector.TransformBy(Matrix3d.Rotation((myBlockRef.Rotation + Math.PI), curUCS.Zaxis, myBlockRef.Position))
                                                            myInPointVector = myInPointVector.TransformBy(Matrix3d.Scaling((myBlockRef.ScaleFactors.X), myBlockRef.Position))

                                                            ''Move the block reference
                                                            myBlockRef.UpgradeOpen()
                                                            myBlockRef.TransformBy(Matrix3d.Displacement(myInPointVector))
                                                        End If
                                                    End If
                                                Next

                                            End If
                                        Case Else
                                            myTrans.Abort()
                                            GoTo CANCEL
                                    End Select
                                End If
                                ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                                ''*********************************************************************************************************
                                ''Change the block rotation:
                                If (My.Settings.BlockRealignMode Or WJL1BlockTools.BlockTools.RealignOpts.Rotation) = My.Settings.BlockRealignMode Then
                                    ''Get the current rotation for the block ref
                                    Dim BlockRotationCurrent As Double = BlockRef.Rotation

                                    ''Get the desired scale from the command line (as a double)
                                    Dim GetAngleMessage As String = vbLf & "Enter new rotation for block '" & BlockRef.Name & "'"
                                    Dim GetAngleOpts As New PromptAngleOptions(GetAngleMessage)
                                    GetAngleOpts.BasePoint = BlockRef.Position
                                    GetAngleOpts.UseBasePoint = True
                                    GetAngleOpts.DefaultValue = BlockRotationCurrent
                                    GetAngleOpts.UseDefaultValue = True
                                    GetAngleOpts.UseAngleBase = True
                                    Dim GetAngleResult As PromptDoubleResult = myEd.GetAngle(GetAngleOpts)

                                    Select Case GetAngleResult.Status
                                        Case PromptStatus.OK
                                            ''If the new rotation differs from the old rotation (waste of time otherwise)
                                            If GetAngleResult.Value <> BlockRotationCurrent Then
                                                Dim RotationAngle As Double = GetAngleResult.Value

                                                ''Redefine the block, rotating everything to match the new rotation
                                                BlockDefinition.UpgradeOpen()
                                                For Each myObjID As ObjectId In BlockDefinition
                                                    ''Scale each object in the block definition
                                                    Dim myEnt As Entity = myObjID.GetObject(OpenMode.ForWrite)
                                                    myEnt.TransformBy(Matrix3d.Rotation((-1 * GetAngleResult.Value), curUCS.Zaxis, New Point3d(0, 0, 0)))
                                                Next

                                                ''Get ModelSpace 
                                                Dim myBT As BlockTable = myDB.BlockTableId.GetObject(OpenMode.ForRead)
                                                Dim myModelSpace As BlockTableRecord = myBT(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForRead)

                                                'For Each Object in ModelSpace
                                                For Each myObjID As ObjectId In myModelSpace
                                                    Dim myEnt As Entity = myObjID.GetObject(OpenMode.ForRead)


                                                    ''Check if the entity is a block refernce
                                                    If TypeOf myEnt Is BlockReference Then
                                                        Dim myBlockRef As BlockReference = myEnt

                                                        ''Check that it's the right block
                                                        If myBlockRef.Name = BlockDefinition.Name Then
                                                            ''Scale the block reference
                                                            myBlockRef.UpgradeOpen()
                                                            myBlockRef.Rotation = myBlockRef.Rotation + RotationAngle
                                                        End If
                                                    End If
                                                Next

                                            End If
                                        Case Else
                                            myTrans.Abort()
                                            GoTo CANCEL
                                    End Select
                                End If
                                ''*********************************************************************************************************

                                ''#########################################################################################################
                                ''Change the block scale:
                                If (My.Settings.BlockRealignMode Or WJL1BlockTools.BlockTools.RealignOpts.Scale) = My.Settings.BlockRealignMode Then
                                    ''Get the current X scale for the block ref
                                    Dim BlockScaleXCurrent As Double = BlockRef.ScaleFactors.X

                                    ''Get the desired scale from the command line (as a double)
                                    Dim GetDoubleMessage As String = vbLf & "Enter new (x) scale for block '" & BlockRef.Name & "'"
                                    Dim GetDoubleOpts As New PromptDoubleOptions(GetDoubleMessage)
                                    GetDoubleOpts.AllowZero = False
                                    GetDoubleOpts.AllowNegative = False
                                    GetDoubleOpts.DefaultValue = BlockScaleXCurrent
                                    GetDoubleOpts.UseDefaultValue = True
                                    Dim GetDoubleResult As PromptDoubleResult = myEd.GetDouble(GetDoubleOpts)

                                    Select Case GetDoubleResult.Status
                                        Case PromptStatus.OK
                                            ''If the new scale differs from the old scale (waste of time otherwise)
                                            If GetDoubleResult.Value <> BlockScaleXCurrent Then
                                                ''Calculate the scale factors
                                                Dim BlockScaleXNew As Double = GetDoubleResult.Value
                                                Dim BlockScaleFactor As Double = BlockScaleXNew / BlockScaleXCurrent

                                                ''Redefine the block, scaling everything up/down to match the new scale
                                                BlockDefinition.UpgradeOpen()
                                                For Each myObjID As ObjectId In BlockDefinition
                                                    ''Scale each object in the block definition
                                                    Dim myEnt As Entity = myObjID.GetObject(OpenMode.ForWrite)
                                                    myEnt.TransformBy(Matrix3d.Scaling((1 / BlockScaleFactor), New Point3d(0, 0, 0)))
                                                Next

                                                ''Get ModelSpace 
                                                Dim myBT As BlockTable = myDB.BlockTableId.GetObject(OpenMode.ForRead)
                                                Dim myModelSpace As BlockTableRecord = myBT(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForRead)

                                                'For Each Object in ModelSpace
                                                For Each myObjID As ObjectId In myModelSpace
                                                    Dim myEnt As Entity = myObjID.GetObject(OpenMode.ForRead)

                                                    ''Check if the entity is a block refernce
                                                    If TypeOf myEnt Is BlockReference Then
                                                        Dim myBlockRef As BlockReference = myEnt

                                                        ''Check that it's the right block
                                                        If myBlockRef.Name = BlockDefinition.Name Then

                                                            ''Calculate the block reference scales
                                                            Dim XScale As Double = myBlockRef.ScaleFactors.X * BlockScaleFactor
                                                            Dim YScale As Double = myBlockRef.ScaleFactors.Y * BlockScaleFactor
                                                            Dim ZScale As Double = myBlockRef.ScaleFactors.Z * BlockScaleFactor

                                                            ''Scale the block reference
                                                            myBlockRef.UpgradeOpen()
                                                            myBlockRef.ScaleFactors = New Geometry.Scale3d(XScale, YScale, ZScale)
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        Case Else
                                            myTrans.Abort()
                                            GoTo CANCEL
                                    End Select
                                End If
                                ''#########################################################################################################
                            End If

                        End If
                        myTrans.Commit()
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                        myEd.WriteMessage(ex.Message)
                        myTrans.Abort()
                    End Try
            End Select
CANCEL:
        End Sub


        Sub SuperExplode()

            Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
            Dim myDB As Database = myDWG.Database
            Dim myEd As Editor = myDWG.Editor

            Dim curUCSMatrix As Matrix3d = myDWG.Editor.CurrentUserCoordinateSystem
            Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d

KEYWORD_ENTERED:
            ''Set the GetEntity options
            Dim GetEntOpts As New PromptEntityOptions("")

            ''Add the Keywords & Message
            GetEntOpts.Message = vbLf & "Select Block to explode: "

            GetEntOpts.AppendKeywordsToMessage = False

            ''Filter so only block references get selected
            GetEntOpts.SetRejectMessage(vbLf & "Selected entity is not a Block.")
            GetEntOpts.AddAllowedClass(GetType(BlockReference), False)

            Dim GetEntResult As PromptEntityResult = myEd.GetEntity(GetEntOpts)

            Select Case GetEntResult.Status
                Case PromptStatus.Keyword
                    Select Case GetEntResult.StringResult
                        Case Else
                            MsgBox("Unhandled keyword: " & GetEntResult.StringResult)
                    End Select
                Case PromptStatus.OK
                    Dim myTrans As Transaction = myDWG.TransactionManager.StartTransaction
                    Try
                        Dim SelectedObject As Object = myTrans.GetObject(GetEntResult.ObjectId, OpenMode.ForRead)
                        Dim SelectedEntity As Entity = CType(SelectedObject, Entity)

                        Dim BlockRef As BlockReference
                        BlockRef = SelectedEntity

                        Dim BlockDefinition As BlockTableRecord = myTrans.GetObject(BlockRef.BlockTableRecord, OpenMode.ForRead)
                        If BlockDefinition.IsFromExternalReference Then
                            myEd.WriteMessage(vbLf & "Selected entity is not a Block.")
                        Else
                            If BlockDefinition.IsDynamicBlock Then
                                myEd.WriteMessage(vbLf & "Selected entity is a Dynamic block.")
                            Else
                                ''Upgrade the BlockDefinition for amending
                                BlockDefinition.UpgradeOpen()

                                ''Store the block definition's "explodable" setting
                                Dim BlockExplodable As Boolean
                                BlockExplodable = BlockDefinition.Explodable

                                ''Set the block definition to "explodable"
                                BlockDefinition.Explodable = True

                                ''Open the Block table for read
                                Dim myBT As BlockTable
                                myBT = myTrans.GetObject(myDB.BlockTableId, OpenMode.ForRead)

                                ''Open the Block table record Model space for write
                                Dim CurrentSpaceBTR As BlockTableRecord
                                CurrentSpaceBTR = TryCast(myTrans.GetObject(myDB.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

                                ''Explode the block reference
                                Dim BlockRefContents As DBObjectCollection = New DBObjectCollection()
                                BlockRef.Explode(BlockRefContents)

                                For Each BlockRefComponant As Entity In BlockRefContents
                                    '' Add the new object to the block table record and the transaction
                                    CurrentSpaceBTR.AppendEntity(BlockRefComponant)
                                    myTrans.AddNewlyCreatedDBObject(BlockRefComponant, True)
                                Next

                                ''Delete the block reference (leaving its exploded contents)
                                BlockRef.UpgradeOpen()
                                BlockRef.Erase()

                                ''Reset block definition to its original "explodable" setting
                                BlockDefinition.Explodable = BlockExplodable
                            End If
                        End If
                        myTrans.Commit()
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                        myEd.WriteMessage(ex.Message)
                        myTrans.Abort()
                    End Try
            End Select
CANCEL:
        End Sub
    End Class
End Namespace

