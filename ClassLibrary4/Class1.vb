Imports System.Diagnostics.PerformanceData
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class Class1
    <CommandMethod("CreateBlockEx")>
    Sub CreateBlockEx()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim opts As PromptSelectionOptions = New PromptSelectionOptions()
        opts.MessageForAdding = "Select entities: "
        Dim selRes As PromptSelectionResult = ed.GetSelection(opts)

        If (selRes.Status <> PromptStatus.OK) Then
            Return
        End If

        Dim ids As ObjectId() = selRes.Value.GetObjectIds()
        Dim blockId As ObjectId = ObjectId.Null

        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim bt As BlockTable = TryCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = TryCast(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead),
                                                  BlockTableRecord)
            Dim counter As Integer = CounterSetting(bt, db)

            If (Not bt.Has("valve" + counter.ToString())) Then


                bt.UpgradeOpen()

                Dim record As BlockTableRecord = New BlockTableRecord()
                record.Name = "valve" + counter.ToString()

                blockId = bt.Add(record)
                trans.AddNewlyCreatedDBObject(record, True)

                If blockId <> ObjectId.Null Then
                    Using acBlkRef As New BlockReference(New Point3d(0, 0, 0), blockId)
                        Dim acCurSpaceBlkTblRec As BlockTableRecord
                        acCurSpaceBlkTblRec = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
                        acCurSpaceBlkTblRec.AppendEntity(acBlkRef)
                        trans.AddNewlyCreatedDBObject(acBlkRef, True)
                    End Using
                End If
            End If
            trans.Commit()

        End Using

        Dim collection As ObjectIdCollection = New ObjectIdCollection(ids)
        Dim mapping As IdMapping = New IdMapping()
        db.DeepCloneObjects(collection, blockId, mapping, True)


        ed.Regen()
        ed.UpdateScreen()
        doc.SendStringToExecute(vbCr + "UPDATETHUMBSNOW" + vbCr, True, False, False)
    End Sub

    <CommandMethod("BlockCre")>
    Sub BlockCCre()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim count as String = ""
        
        Dim opts As PromptSelectionOptions = New PromptSelectionOptions()
        opts.MessageForAdding = "Select entities: "
        Dim selRes As PromptSelectionResult = ed.GetSelection(opts)

        If (selRes.Status <> PromptStatus.OK) Then
            Return
        End If

        Dim ids As ObjectId() = selRes.Value.GetObjectIds()
        Dim blockId As ObjectId = ObjectId.Null

        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim bt As BlockTable = TryCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = TryCast(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead),
                                                  BlockTableRecord)
            Dim counter As Integer = CounterSetting(bt, db)

            If (Not bt.Has("valve" + counter.ToString())) Then
                bt.UpgradeOpen()
                Dim record As BlockTableRecord = New BlockTableRecord()
                record.Name = "valve" + counter.ToString()
                count = "valve" + counter.ToString()
                blockId = bt.Add(record)
                trans.AddNewlyCreatedDBObject(record, True)
                If blockId <> ObjectId.Null Then
                    Using acBlkRef As New BlockReference(New Point3d(0, 0, 0), blockId)
                        Dim acCurSpaceBlkTblRec As BlockTableRecord
                        acCurSpaceBlkTblRec = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
                        acCurSpaceBlkTblRec.AppendEntity(acBlkRef)
                        trans.AddNewlyCreatedDBObject(acBlkRef, True)
                    End Using
                End If
            End If

            trans.Commit()

        End Using
        
        Dim collection As ObjectIdCollection = New ObjectIdCollection(ids)
        Dim mapping As IdMapping = New IdMapping()
        db.DeepCloneObjects(collection, blockId, mapping, True)

      
        ed.Regen()
        ed.UpdateScreen()
        doc.SendStringToExecute(vbCr + "UPDATETHUMBSNOW" + vbCr, True, False, False)
        
  
    End Sub
    
    Public Sub AddingAttributeToABlock()
        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            If Not acBlkTbl.Has("CircleBlockWithAttributes") Then
                Using acBlkTblRec As New BlockTableRecord
                    acBlkTblRec.Name = "CircleBlockWithAttributes"

                    ' Set the insertion point for the block
                    acBlkTblRec.Origin = New Point3d(0, 0, 0)

                    ' Add a circle to the block
                    Using acCirc As New Circle
                        acCirc.Center = New Point3d(0, 0, 0)
                        acCirc.Radius = 2

                        acBlkTblRec.AppendEntity(acCirc)

                        ' Add an attribute definition to the block
                        Using acAttDef As New AttributeDefinition
                            acAttDef.Position = New Point3d(0, 0, 0)
                            acAttDef.Verifiable = True
                            acAttDef.Prompt = "Door #: "
                            acAttDef.Tag = "Door#"
                            acAttDef.TextString = "DXX"
                            acAttDef.Height = 1
                            acAttDef.Justify = AttachmentPoint.MiddleCenter
                            acBlkTblRec.AppendEntity(acAttDef)

                            acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
                            acBlkTbl.Add(acBlkTblRec)
                            acTrans.AddNewlyCreatedDBObject(acBlkTblRec, True)
                        End Using
                    End Using
                End Using
            End If

            ' Save the new object to the database
            acTrans.Commit()

            ' Dispose of the transaction
        End Using
    End Sub
    
    <CommandMethod("INJECTOR", CommandFlags.Session)>
     Sub Injector()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim acdb As Database = doc.Database
        Dim opts As New PromptEntityOptions(vbNewLine & "Select Block:")
        Dim res As PromptEntityResult = ed.GetEntity(opts)
        If res.Status <> PromptStatus.OK Then Exit Sub
        Dim id As ObjectId = res.ObjectId
        Using doc.LockDocument
            Using tr As Transaction = doc.Database.TransactionManager.StartTransaction
                Dim blk As BlockReference = tr.GetObject(id, OpenMode.ForRead)
                Dim blkName As String = blk.Name.ToUpper()
                Dim bt As BlockTable = tr.GetObject(acdb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(bt(blkName), OpenMode.ForWrite)
                If btr.Name.ToUpper() = blkName Then
                    btr.UpgradeOpen()
                    Dim brefIds As ObjectIdCollection = btr.GetBlockReferenceIds(False, True)
                    Dim stropts As New PromptStringOptions(vbNewLine & "Attribute Name:")
                    Dim strres As PromptResult = ed.GetString(stropts)
                    If strres.Status <> PromptStatus.OK OrElse strres.StringResult = "CANCEL" Then Exit Sub
                    Dim attName As String = strres.StringResult
                    Dim attDef As New AttributeDefinition()
                    attDef.Verifiable = True
                    attDef.Tag = attName
                    attDef.Prompt = attName

                    attDef.Justify = AttachmentPoint.MiddleCenter
                    attDef.Invisible = True
                    attDef.Height = 3
                    btr.AppendEntity(attDef)
                    tr.AddNewlyCreatedDBObject(attDef, True)

                    btr.DowngradeOpen()
                    ed.UpdateScreen()
                    ed.WriteMessage(vbNewLine & "Updating existing block references.")
                    For Each objid As ObjectId In brefIds
                        Dim bref As BlockReference = tr.GetObject(objid, OpenMode.ForWrite, False, True)
                        bref.RecordGraphicsModified(True)
                        bref.UpgradeOpen()
                    Next

                End If
                tr.Commit()
            End Using
        End Using
    End Sub
    
     Sub InjectorBlock()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim acdb As Database = doc.Database
        Dim opts As New PromptEntityOptions(vbNewLine & "Select Block:")
        Dim res As PromptEntityResult = ed.GetEntity(opts)
         
         
        If res.Status <> PromptStatus.OK Then Exit Sub
        Dim id As ObjectId = res.ObjectId
        Using doc.LockDocument
            Using tr As Transaction = doc.Database.TransactionManager.StartTransaction
                Dim blk As BlockReference = tr.GetObject(id, OpenMode.ForRead)
                Dim blkName As String = blk.Name.ToUpper()
                Dim bt As BlockTable = tr.GetObject(acdb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(bt(blkName), OpenMode.ForWrite)
                If btr.Name.ToUpper() = blkName Then
                    btr.UpgradeOpen()
                    Dim brefIds As ObjectIdCollection = btr.GetBlockReferenceIds(False, True)
                    Dim stropts As New PromptStringOptions(vbNewLine & "Attribute Name:")
                    Dim strres As PromptResult = ed.GetString(stropts)
                    If strres.Status <> PromptStatus.OK OrElse strres.StringResult = "CANCEL" Then Exit Sub
                    Dim attName As String = strres.StringResult
                    Dim attDef As New AttributeDefinition()
                    attDef.Verifiable = True
                    attDef.Tag = attName
                    attDef.Prompt = attName

                    attDef.Justify = AttachmentPoint.MiddleCenter
                    attDef.Invisible = True
                    attDef.Height = 3
                    btr.AppendEntity(attDef)
                    tr.AddNewlyCreatedDBObject(attDef, True)

                    btr.DowngradeOpen()
                    ed.UpdateScreen()
                    ed.WriteMessage(vbNewLine & "Updating existing block references.")
                    For Each objid As ObjectId In brefIds
                        Dim bref As BlockReference = tr.GetObject(objid, OpenMode.ForWrite, False, True)
                        bref.RecordGraphicsModified(True)
                        bref.UpgradeOpen()
                    Next

                End If
                tr.Commit()
            End Using
        End Using
    End Sub
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
   
    
    Sub AddBlockTest(blockName as String)
        Dim db as Database = Application.DocumentManager.MdiActiveDocument.Database
        
        Using myt as Transaction = db.TransactionManager.StartTransaction()
            Dim bt as BlockTable = TryCast(db.BlockTableId.GetObject(OpenMode.ForRead), BlockTable)
            Dim blockDef as BlockTableRecord  = TryCast(bt(blockName).GetObject(OpenMode.ForRead), BlockTableRecord)
            Dim ms As BlockTableRecord = TryCast(bt(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite), BlockTableRecord)
            
            Dim point As Point3d = New Point3d(0,0,0)
            Using blockRef as BlockReference = New BlockReference(point, blockDef.ObjectId)
                ms.AppendEntity(blockRef)
                myt.AddNewlyCreatedDBObject(blockRef, True)
                
                For Each id As ObjectId in blockDef
                    Dim obj As DBObject = id.GetObject(OpenMode.ForRead)
                    Dim attDef as AttributeDefinition = DirectCast(obj, AttributeDefinition)
                    
                    If((attDef IsNot Nothing) And (Not attDef.Constant))
                        Using attRef As AttributeReference = new AttributeReference()
                            attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform)
                            attRef.TextString = "Hello world"
                            blockRef.AttributeCollection.AppendAttribute(attRef)
                            myt.AddNewlyCreatedDBObject(attRef, True)
                        End Using
                    End If
                Next
            End Using
            myt.Commit()
        End Using
    End Sub
    
    
    Sub Attribute(ByVal BlokRefId As ObjectId)
        Dim doc = Application.DocumentManager.MdiActiveDocument
        Dim dwg = doc.Database

        Using doc.LockDocument()
            Using transactie = doc.TransactionManager.StartTransaction()
                Try
                    Dim Ref As BlockReference
                    Ref = transactie.GetObject(BlokRefId, OpenMode.ForWrite)
                    Dim a = Ref.Name

                    Dim BlokDefinities As BlockTable
                    BlokDefinities = transactie.GetObject(dwg.BlockTableId, OpenMode.ForRead)
                    Dim Blokdefid = BlokDefinities(Ref.Name)
                    Dim BlokDefinitie As BlockTableRecord
                    BlokDefinitie = transactie.GetObject(Blokdefid, OpenMode.ForRead)

                    Dim AttRefIdColl = Ref.AttributeCollection

                    For Each elementId In BlokDefinitie
                        Dim Element As Entity
                        Element = transactie.GetObject(elementId, OpenMode.ForRead)

                        If TypeOf Element Is AttributeDefinition Then
                            Dim attribuutdefinitie = CType(Element, AttributeDefinition)
                            Dim attribuutreferentie As New AttributeReference
                            attribuutreferentie.SetAttributeFromBlock(attribuutdefinitie, Ref.BlockTransform)
                            AttRefIdColl.AppendAttribute(attribuutreferentie)
                            transactie.AddNewlyCreatedDBObject(attribuutreferentie, True)
                        End If
                    Next

                    transactie.Commit()
                Catch ex As Exception
                    MsgBox("Er ging iets fout: " & vbCrLf & ex.Message)
                End Try
            End Using
        End Using
    End Sub
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    <CommandMethod("CreateBlockEx1")>
    Sub CreateBlockEx1()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim opts As PromptSelectionOptions = New PromptSelectionOptions()
        opts.MessageForAdding = "Select entities: "
        Dim selRes As PromptSelectionResult = ed.GetSelection(opts)

        If (selRes.Status <> PromptStatus.OK) Then
            Return
        End If

        Dim ids As ObjectId() = selRes.Value.GetObjectIds()
        Dim blockId As ObjectId = ObjectId.Null

        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim bt As BlockTable = TryCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = TryCast(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead),
                                                  BlockTableRecord)
            Dim counter As Integer = CounterSetting(bt, db)

            If (Not bt.Has("valve" + counter.ToString())) Then

                bt.UpgradeOpen()

                Dim record As BlockTableRecord = New BlockTableRecord()
                record.Name = "valve" + counter.ToString()

                blockId = bt.Add(record)
                trans.AddNewlyCreatedDBObject(record, True)

                If blockId <> ObjectId.Null Then
                    Using acBlkRef As New BlockReference(New Point3d(0, 0, 0), blockId)

                        Dim acCurSpaceBlkTblRec As BlockTableRecord
                        acCurSpaceBlkTblRec = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
                        acCurSpaceBlkTblRec.AppendEntity(acBlkRef)
                        trans.AddNewlyCreatedDBObject(acCurSpaceBlkTblRec, True)
                    End Using


                End If


            End If
            trans.Commit()

        End Using

        Dim collection As ObjectIdCollection = New ObjectIdCollection(ids)
        Dim mapping As IdMapping = New IdMapping()
        db.DeepCloneObjects(collection, blockId, mapping, True)


        ed.Regen()
        ed.UpdateScreen()
        doc.SendStringToExecute(vbCr + "UPDATETHUMBSNOW" + vbCr, True, False, False)
    End Sub


    Private Function CounterSetting(bt As BlockTable, db As Database) As Integer
        Dim counter As Integer = 0
        Using trans As Transaction = db.TransactionManager.StartTransaction
            For Each btrId As ObjectId In bt
                Dim btr As BlockTableRecord = trans.GetObject(btrId, OpenMode.ForRead)

                If btr.Name.Contains("valve") Then
                    counter = counter + 1
                End If
            Next
        End Using

        Return counter
    End Function


   
End Class