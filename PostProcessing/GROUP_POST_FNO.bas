Sub Main
    Dim App As femap.model
    Set App = feFemap()

                Dim ouSets As femap.Set
                Dim ouVecs As femap.Set

                Dim feGroup As femap.Group
                Set feGroup = App.feGroup

                Dim fName As String
                Dim neu_Name As String
                Dim fno_Name As String
                Dim createdGroup As Boolean

                createdGroup = False

                rc = feGroup.SelectID( "Select Group to Process" )
                If rc <> -1 Then
                                GoTo Jump_Out
                End If

                rc = feGroup.AddRelated()
                rc = feGroup.Put( feGroup.NextEmptyID )
                createdGroup = True

                fName = feGroup.title

                neu_Name = fName+".neu"
                fno_Name = fName+".fno"

                rc = App.feFileWriteNeutral2( 0, neu_Name, False, True, False, False, False, False, False, False, False, False, 20, 0.0, feGroup.ID )

                rc = App.feSelectOutput( "Select Output to Export", 0, 0, FOC_ANY, 0, True, ouSets, ouVecs )

                If rc <> -1 Then
                                MsgBox( "Error Getting Output to Export", vbOkOnly, "Sub-Model and Post" )
                                GoTo Jump_Out
                End If

                rc = App.feFileWriteFNO( ouSets.ID, ouVecs.ID, feGroup.ID, fno_Name )

                Dim modID As Huge_

                App.feAppGetModel( modID )
                rc = App.feFileNew()
                rc = App.feFileReadNeutral( 0, neu_Name, True, True, False, True, 0 )
                rc = App.feFileAttachResults( FAP_NE_NASTRAN, fno_Name, False )
                App.feAppSetModel( modID )

Jump_Out:
If createdGroup Then
                rc = feGroup.Delete( feGroup.ID )
End If

End Sub
