Imports System.Diagnostics.PerformanceData
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class Class1

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

      
 
        valveSpecInjection(count)
        
        doc.SendStringToExecute("._ATTSYNC _N " + count , True, False, False)
        
        ed.Regen()
        ed.UpdateScreen()
        doc.SendStringToExecute(vbCr + "UPDATETHUMBSNOW" + vbCr, True, False, False)
        
    End Sub
    
    Sub valveSpecInjection(blockName as String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        
        Using doc.LockDocument
            Using tr As Transaction = doc.Database.TransactionManager.StartTransaction
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(bt(blockName), OpenMode.ForWrite)
                Dim blkName As String = btr.Name.ToUpper()
                If btr.Name.ToUpper() = blkName Then
                    btr.UpgradeOpen()
                    Dim brefIds As ObjectIdCollection = btr.GetBlockReferenceIds(False, True)
                    
                    Dim inputString As String = ""
                    Dim attValveNo As New AttributeDefinition()
                    With attValveNo
                        inputString = "ValveNo"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attValveNo)
                    tr.AddNewlyCreatedDBObject(attValveNo, True)
                    
                    Dim attValvePurpose As New AttributeDefinition()
                    With attValvePurpose
                        inputString = "PURPOSE"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attValvePurpose)
                    tr.AddNewlyCreatedDBObject(attValvePurpose, True)
                    
                    Dim attValveType As New AttributeDefinition()
                    With attValveType
                        inputString = "TYPE"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attValveType)
                    tr.AddNewlyCreatedDBObject(attValveType, True)
                    
                    Dim attQTY As New AttributeDefinition()
                    With attQTY
                        inputString = "QTY"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attQTY)
                    tr.AddNewlyCreatedDBObject(attQTY, True)
                    
                    Dim attNomalDia As New AttributeDefinition()
                    With attNomalDia
                        inputString = "Nom.Dia."
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attNomalDia)
                    tr.AddNewlyCreatedDBObject(attNomalDia, True)
                    
                    Dim materialA As New AttributeDefinition()
                    With materialA
                        inputString = "MaterialA"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(materialA)
                    tr.AddNewlyCreatedDBObject(materialA, True)
                    
                    Dim materialB As New AttributeDefinition()
                    With materialB
                        inputString = "MaterialB"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(materialB)
                    tr.AddNewlyCreatedDBObject(materialB, True)
                    
                    Dim attConnectionType As New AttributeDefinition()
                    With attConnectionType
                        inputString = "Con.Type"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attConnectionType)
                    tr.AddNewlyCreatedDBObject(attConnectionType, True)
                    
                    Dim attConnectionClass As New AttributeDefinition()
                    With attConnectionClass
                        inputString = "Con.Class"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attConnectionClass)
                    tr.AddNewlyCreatedDBObject(attConnectionClass, True)

                    Dim attWorkingPressure As New AttributeDefinition()
                    With attWorkingPressure
                        inputString = "W.P"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attWorkingPressure)
                    tr.AddNewlyCreatedDBObject(attWorkingPressure, True)
                    
                    Dim attWorkingTemperature As New AttributeDefinition()
                    With attWorkingTemperature
                        inputString = "W.T."
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attWorkingTemperature)
                    tr.AddNewlyCreatedDBObject(attWorkingTemperature, True)

                    Dim attTestTemperature As New AttributeDefinition()
                    With attTestTemperature
                        inputString = "T.P."
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attTestTemperature)
                    tr.AddNewlyCreatedDBObject(attTestTemperature, True)


                    Dim attAaboveB As New AttributeDefinition()
                    With attAaboveB
                        inputString = "A/B"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attAaboveB)
                    tr.AddNewlyCreatedDBObject(attAaboveB, True)

                    Dim attPositionOfAccess As New AttributeDefinition()
                    With attPositionOfAccess
                        inputString = "Location"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attPositionOfAccess)
                    tr.AddNewlyCreatedDBObject(attPositionOfAccess, True)
                    
                    Dim attMAKER As New AttributeDefinition()
                    With attMAKER
                        inputString = "Maker"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attMAKER)
                    tr.AddNewlyCreatedDBObject(attMAKER, True)
                    
                    Dim attModel As New AttributeDefinition()
                    With attModel
                        inputString = "MODEL"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attModel)
                    tr.AddNewlyCreatedDBObject(attModel, True)

                    Dim attPorNo As New AttributeDefinition()
                    With attPorNo
                        inputString = "POR NO"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attPorNo)
                    tr.AddNewlyCreatedDBObject(attPorNo, True)

                    Dim attGSI As New AttributeDefinition()
                    With attGSI
                        inputString = "GSI"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attGSI)
                    tr.AddNewlyCreatedDBObject(attGSI, True)
                    
                    Dim attActuator As New AttributeDefinition()
                    With attActuator
                        inputString = "Actuator"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attActuator)
                    tr.AddNewlyCreatedDBObject(attActuator, True)
                    
                    Dim attBlock As New AttributeDefinition()
                    With attBlock
                        inputString = "BLOCK"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attBlock)
                    tr.AddNewlyCreatedDBObject(attBlock, True)
                    
                    Dim attEVENT As New AttributeDefinition()
                    With attEVENT
                        inputString = "EVENT"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attEVENT)
                    tr.AddNewlyCreatedDBObject(attEVENT, True)
                    
                    Dim attPersonInCharge As New AttributeDefinition()
                    With attPersonInCharge
                        inputString = "PERSON"
                        .Verifiable = False
                        .Tag = inputString
                        .Prompt = inputString
                        .TextString = "-"
                        .Invisible = True
                    End With
                    btr.AppendEntity(attPersonInCharge)
                    tr.AddNewlyCreatedDBObject(attPersonInCharge, True)

                    
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