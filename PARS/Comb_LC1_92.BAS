Sub Main ()
Dim rc As Long
Dim lastOS As Long


Dim femap As Object
Set femap = GetObject(, "femap.model")


Dim Snum As Long
Dim Title As String
Dim Pozn As String

'promenne pro kombinaci
Dim Radku As Variant
Dim ov (100)  As Variant
Dim ply (100) As Variant
Dim info (100) As Variant
Dim gr (100) As Variant
Dim OsProc (100) As Variant
Dim fact  (100) As Variant
Dim vyslOS (100) As Variant

Dim os As Object
Dim aa As Long



'nacti jmeno souboru
Dim fName As String

rc = femap.feFileGetName("Otevri soubor LC kombinaci", "Linear Comb. File","*.cmb", True, fName)
If rc=0 Then GoTo Konec

'print rc

' vytor objekt file
Dim f As Object
Set f = femap.feRead

rc = f.Open(fName,100) 'otevre soubor
'rc = f.ShowStatus(True)

rc = f.SetFreeFormat() 'volny format, oddelovac carka

Do 'do konce souboru

   rc = f.Read()     '1 radek
   'print f.line
   If chyba(rc) Then GoTo Konec
   Snum=f.IntField( 2, 0)
   'print Snum


   rc = f.Read()     '2 radek
   If chyba(rc) Then GoTo Konec
   Title = f.Line

   rc = f.Read()     '3 radek
   If chyba(rc) Then GoTo Konec
   Pozn = f.Line

     
   Radku = 0
   Do 
     rc = f.Read()
     If chyba(rc) Then GoTo Konec
     OsProc(Radku)=f.IntField( 1, 1)  'kombinuj stav cislo

     fact(Radku)=f.RealField( 2, 0.0) ' s faktorem

     'Print OsProc(Radku), fact(Radku)

     Radku=Radku+1

   Loop Until OsProc(Radku-1)=-999 'do konce kombinace
   rc = f.Read() 'nacte prazdny radek

	Set OsProc(Radku-1) = Empty
	Set fact(Radku-1) = Empty

   Set os = femap.feOutputSet
   rc = os.Delete(Snum) 'vymaze Output set s danym cislem (pokud existuje)


   'Nadi posledni set
   'rc=os.Last() 'posledni set
   'aa=os.ID ' a jeho cislo
   'Snum=aa+1 'prvni volne cislo setu

   'Initializuje Output Set a ulozi ho

   os.title = Title
   os.notes = Pozn

   os.program = 4 'Nastran
   os.analysis = 1 'linearni statika
   os.Value =0.0

   os.Put (Snum) ' ulozi set pod cislem Snum 


   'Kombinace zat stavu


   vyslOS(0)=Snum


   rc= femap.feOutputProcess( Radku-1 , True, 2, gr, OsProc, gr , vyslOS ,gr , gr , fact, ov, ply, info )
   If rc<>-1 Then MsgBox ("Chyba pri vytvareni kombinace")
   'Print rc

Loop Until f.AtEOF() 'do konce souboru

rc = f.Close()

Konec:
' Print "konec"
End Sub

Function chyba(rc As Long)
   If rc<>-1 Then MsgBox ("Chyba pri nacitani souboru")
   chyba = 0
End Function
