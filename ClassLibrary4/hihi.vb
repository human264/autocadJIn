Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class hihi
     <CommandMethod("AddingAttributeToABlock")> _
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

                             acBlkTbl.UpgradeOpen()
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
     
     
     
End Class