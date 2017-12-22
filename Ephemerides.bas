REM  *****  BASIC  *****

sub codes_format_des_cellules
Dim NumberFormats As Object
Dim NumberFormatString As String
Dim NumberFormatId As Long
Dim LocalSettings As New com.sun.star.lang.Locale

Doc = ThisComponent
Sheet = Doc.Sheets.getByName("thème")
Cell = Sheet.getCellByPosition(2,2)
 
Cell.Value = 2
'Cell.CellBackColor = RGB(0, 100, 0)

LocalSettings.Language = "fr" '"en"
LocalSettings.Country = "fr" '"us"
 
NumberFormats = Doc.NumberFormats
'NumberFormatString ="JJ/MM/AAAA" ' "#"  'pas de décimales : "#,##0" = 3 décimales
NumberFormatString = "00"

NumberFormatId = NumberFormats.queryKey(NumberFormatString, LocalSettings, True)
If NumberFormatId = -1 Then
   NumberFormatId = NumberFormats.addNew(NumberFormatString, LocalSettings)
End If
 
MsgBox NumberFormatId
Cell.NumberFormat = NumberFormatId '10030: date jj/mm/aaaa, 10001 nombre sans décimale, 10002 2 décimales,10107 : 01 au lieu de 1
end sub

sub filtre_feuille_exemple
Dim oFilterDesc ' Filter descriptor.
Dim oFields(0) As New com.sun.star.sheet.TableFilterField
Doc = thisComponent
Sheet = Doc.Sheets.getByName("éphémérides")
Sheet = Doc.Sheets.getByName("transits2")
	'dernière ligne ?
	Curs = Sheet.createCursor
	Curs.gotoEndOfUsedArea(True)
	lastrow = Curs.Rows.Count 
'effacement filtre
oFilterDesc = Sheet.createFilterDescriptor(True)
Sheet.filter(oFilterDesc)
oFilterDesc = Sheet.createFilterDescriptor(True)
'définition des filtres
With oFields(0)
.Field = 6 'colonne
.IsNumeric = false
.StringValue = "phase 1" 'ou numericvalue
.Operator = com.sun.star.sheet.FilterOperator.EQUAL
End With

'exécution
oFilterDesc.setFilterFields(oFields())
Sheet.filter(oFilterDesc)
Cell = Sheet.getCellrangeByPosition(1,0,1,lastrow-1) : cell.cellBackColor = RGB(255, 255, 155)
	'dernière ligne ?
	Curs = Sheet.createCursor
	Curs.gotoEndOfUsedArea(True)
	lastrow = Curs.Rows.Count 
'effacement filtre
oFilterDesc = Sheet.createFilterDescriptor(True)
Sheet.filter(oFilterDesc)
end sub

' test sur format de nombres
sub test
dim chaine as string
dim annee(0 to 1000) as string
dim mois(0 to 100) as string
dim an(0 to 100) as string
dim annum(0 to 100) as long
dim d(0 to 2555) as double
dim repere as integer
Dim Dummy() 
Dim Url As String
dim num as long
Dim CellRangeAddress As New com.sun.star.table.CellRangeAddress
Dim CellAddress As New com.sun.star.table.CellAddress
Dim oFA 'as New com.sun.star.sheet.FunctionAccess
dim aa,bb,cc,dd
dim ss '(0 to 15, 0 to 15) as double
dim tt(0 to 15,0) as integer ' as double
dim th(0 to 30) as double
dim arrondi
dim theme
dim z as double
dim xx(0 to 255,0) as integer
dim cell1,cell2
dim zz(359 to 600)
dim a as double
Dim oSheet
Dim oFilterDesc ' Filter descriptor.
Dim oFields(2) As New com.sun.star.sheet.TableFilterField
Dim aSortFields(0) As New com.sun.star.util.SortField 'pour trier colonnes
Dim aSortDesc(0) As New com.sun.star.beans.PropertyValue 'pour trier colonnes
dim ess%
dim yy(0 to 360)as integer
dim yy2(0 to 360)as integer
dim angle
dim jour
'oFA = oManager.createInstance("com.sun.star.sheet.FunctionAccess")
oFA = createUnoService( "com.sun.star.sheet.FunctionAccess" )
aa= oFA.callFunction("Max",array(10,1523))
car=int(11/3)
car=max(1,2)
angle=379 mod 360
for i=359 to 500
zz(i)=i mod 360
next i

'abc="date"
'a=abc&"value"(now)
'z=today()'format(now,dd/mm/yyyy)
car=int(1439/60)
car1=1438 mod 60
msgbox zz' now'today()
a=datevalue(now)
abc="~/Documents/"
bcd=mid$(abc,5)
for i=13 to 24
yy(i)=i mod 13
next i
oFA = createUnoService( "com.sun.star.sheet.FunctionAccess" )
angle=0
for i=1 to 36
	'conversion en radians
	car=oFA.callFunction("Radians",array(angle))
	'sinus de l'angle
	sinus=oFA.callFunction("Sin",array(car))
	angle=angle+10
	next i
longitude=42
car=oFA.callFunction("Sqrt",array(longitude))
'car=SQRT(16)
ess=44.999 mod 15
for i= 0 to 359
yy(i)=int(i / 15)
yy2(i)=int(i / 3)
next i
ess =10/30
ess=3/2
for i=0 to 32
xx(i,0)=i mod 2
next i
Doc = thisComponent
Sheet = Doc.Sheets.getByName("éphémérides")
car=sheet("temp")
Sheet = Doc.Sheets.getByName("transits2")
zz=array("phase 1","phase 2","phase 3")
oFilterDesc = Sheet.createFilterDescriptor(True)
for i=0 to 2
With oFields(i)
.Field = 6
.IsNumeric = false
.StringValue = zz(i)
.Operator = com.sun.star.sheet.FilterOperator.EQUAL
end with
next i
oFilterDesc.setFilterFields(oFields())
oFilterDesc.CopyOutputData = True
'oFilterDesc.ContainsHeader = True
Sheet.filter(oFilterDesc)

'effacement filtre
oFilterDesc = Sheet.createFilterDescriptor(True)
Sheet.filter(oFilterDesc)

call definitions
'call theme_natal

'oFA = createUnoService( "com.sun.star.sheet.FunctionAccess" )
oManager = GetProcessServiceManager()
'oDesk = oManager.createInstance("com.sun.star.frame.Desktop")

Doc = thisComponent
Sheet = Doc.Sheets.getByName("éphémérides")
car=thisComponent.getSheets.getCount

cell=Sheet.getCellRangeByPosition(0, 1,1, 16)

aa=cell.getdataArray
cell=Sheet.getCellRangeByPosition(31, 1,31, 16)
aa=cell.getdataArray
longitude=192.245
car=oFA.callFunction("Match",array(longitude,plan(3),1))
car=oFA.callFunction(,array(longitude,plan(3),1))
abc=asp(3)(car-1)(0)

'car=dcount(B1:E15; 0; AB1:AE2)
car=oFA.callFunction("DCount",array(cell1,0,cell2))
'car=oFA.callFunction("DGet",array(cell1,1,cell2))

car1=oFA.callFunction("Match",array(246,cell1,0))


car=ubound(tt,2)
car=oFA.callFunction("Match",array(10, tt(),0))
'call theme_natal
longitude= 261.1
abc=Soleil
			' longitude=oFA.callFunction("Round",array(longitude,2))
car=oFA.callFunction("Match",array(longitude,abc,0))
			 
			 
cell=Sheet.getCellRangeByPosition(1, 1,16, 16)
theme=cell.getdata
'for i=1 to 15
cell=Sheet.getCellRangeByPosition(1, 1,1, 16)
soleil=cell.getdata
cell=Sheet.getCellRangeByPosition(2, 1,2, 16)
lune=cell.getdata
'next i
longitude=66.1
c
bb=oFA.callFunction("Round",array(soleil,2))
cc=oFA.callFunction("Round",array(soleil,-1))
car=oFA.callFunction("Match",array(longitude,bb,0))

end sub

Sub appel_fonction
Dim oFA
dim ss
dim aa,bb,cc,dd 'ne pas mettre "as double" ou autre sinon erreur "variable objet non définie"
oFA = createUnoService( "com.sun.star.sheet.FunctionAccess" )
'à tester...
svcfa = createUnoService( "com.sun.star.sheet.FunctionAccess" )
Z = svcfa.callFunction("com.sun.star.sheet.addin.Analysis.getComplex",Array(5,2,"i"))
'autre
'oSelection.setValue(oFunction.callFunction("NOW", Array()))

'optimisation ?
	' Round the value to the given number of places after the decimal.
	'Function Round(value, decimalPlaces) 
	  ' InitRound() 
	   'Dim args( 1 to 2 ) As Variant 
	   'args(1) = value 
	  ' args(2) = decimalPlaces 
	  ' Round = oFunction.callFunction( "round", args() ) 
  'après optimisation
	 ' Function Round2(value, decimalPlaces) as double
   ' Round2 = Int(value * 10 ^ decimalPlaces + 0.5) / 10 ^ decimalPlaces
   'Dim multiplier as double, bigValue as double
   'multiplier = 10 ^ decimalPlaces
   'bigValue = (value * multiplier) + 0.5
   'Round2 = CDbl( CLng(bigValue) ) / multiplier  
'	End Function

	'autre façon de déclarer
	'oManager = GetProcessServiceManager()
	'oDesk = oManager.createInstance("com.sun.star.frame.Desktop")
	'oFA = oManager.createInstance("com.sun.star.sheet.FunctionAccess")

Doc = ThisComponent
Sheet = Doc.Sheets.getByName("éphémérides")
cell=Sheet.getCellRangeByPosition(1, 1,1, 200)
ss=cell.getdata
car1=ubound(ss)
car=ubound(ss)
aa= oFA.callFunction("Max",array(ss,1523))
bb=oFA.callFunction("Round",array(ss,2))
cc=oFA.callFunction("Round",array(ss,-1))
car=oFA.callFunction("Match",array(290.31,bb,0))

for i=1 to 15
cell=Sheet.getCellRangeByPosition(i, 1,i, 16)
ss(i)=cell.getdata
next i
car=oFA.callFunction("Match",array(306.095,ss(1),0))

'database
cell1=Sheet.getCellRangeByPosition(1, 0,4, 16)
cell2=Sheet.getCellRangeByPosition(27, 0,30, 1) 'critères
'car=oFA.callFunction("DCount",array(cell1,0,cell2))

' Calculate min of numbers.
print oFA.callFunction( "MIN", array( 10, 23, 5, 345 ) )
End Sub



public feuille as object
global feuille2 as object
public form_ok%
public commande1, commande2, commande3, commande4, commande5, commande6, commande7, commande8
global Doc As Object
global Sheet as Object
global Cell As Object
Global CellRangeAddress As New com.sun.star.table.CellRangeAddress
global abc as string, bcd as string, cde as string
global annee_progresse%
global angle(0 to 16) as integer
global arc(0 to 24,0 to 1)
global aspect(0 to 16) as string
global bleu as long
global car, car1, car2, car3, car4
global choix
global clic_sortie
global coeff1 as integer, coeff2 as integer, coeff3 as integer
global compte_lignes as long
global couleur(0 to 15) as long
global date_naissance as long
global degres%, minutes%
global hauteur(1)
global h,i,j,k,l,m,n
global index_signe as integer
global jours%
global lastrow as long
global lien(0 to 23) as integer
global longitude as double
global longueur(1)
global matrice(9,11) as string
global mauve as long
global message as string
global num0, num1 as integer
Global oFA as object 'pour utiliser les fonctions min,max,match, etc...
global orbe(0 to 15) as double 'currency  'pour avoir orbe quiconce =0,5 sinon arrondi à 1
global orbe_theme(12,12,16) as double 'planète sujet, agent, orbe
global orbe_transit(0 to 15) as double
global orbedecimal as double ' currency 'permet d'avoir un nombre décimal "en clair" 0,001 pas 1E3
global phase() 'array de lecture phases pour les transits
global plan(0 to 15), asp(0 to 15) 'longitude + aspects du thème natal par ordre croissant de longitudes
global planete(0 to 15) as string
global position_ephe0 as double, position_ephe1 as double, position_ephe2 as double, position_ephe3 as double
Global position_epheref0 as double, position_epheref1 as double, position_epheref2 as double, position_epheref3 as double
global rouge as long
global signe(0 to 11) as string

'note sur array : dim aa(colonne,ligne); si on pévoit d'écrire une ligne de 16 éléments, faire "dim aa(0,0 to 15)" ou si une colonne "dim aa(0 to 15,0)"
sub definitions
car=0
for i= 0 to 23 step 2
lien(i)=car
lien(i+1)=car
'inter-aspects
if i=2 or i=8 or i=14 or i=20 then lien(i+1)=lien(i+1)+15
car=car+30
next i

planete(0)="Soleil" :  planete(1)="Lune" : planete(2)="Mercure" : planete(3)="Vénus" : planete(4)="Mars" : 
planete(5)="Jupiter" : planete(6)="Saturne" : planete(7)="Uranus" : planete(8)="Neptune" : planete(9)="Pluton"
planete(10)="NN" : planete(11)="Lilith" :  planete(12)="AS" : planete(13)="FC" : planete(14)="DS" : planete(15)="MC"

'ne pas enlever les phases, utilisées par les tableaux de transit, planètes et années !
aspect(0)="conjonction " & chr(966) & "1" : aspect(1)="semi-sextile " & chr(966) & "2": aspect(2)="semi-carré i" & chr(966) & "2"
aspect(3)="sextile " & chr(966) & "3" : aspect(4)="carré " & chr(966) & "4" : aspect(5)="trigone " & chr(966) & "5"
aspect(6)="sesqui-carré i" & chr(966) & "5" : aspect(7)="quinconce " & chr(966) & "6" : aspect(8)="opposition " & chr(966) & "7" 
aspect(9)="quinconce " & chr(966) & "8" : aspect(10)="sesqui-carré i" & chr(966) & "8" : aspect(11)="trigone " & chr(966) & "9"
aspect(12)="carré " & chr(966) & "10" : aspect(13)="sextile " & chr(966) & "11" : aspect(14)="semi-carré i" & chr(966) & "11"
aspect(15)="semi-sextile " & chr(966) & "12": aspect(16)="conjonction " & chr(966) & "1"

angle(0)=0 : angle(1)=30: angle(2)=45 : angle(3)=60 : angle(4)=90: angle(5)=120 : angle(6)=135 : angle(7)=150 : angle(8)=180
angle(9)=210 : angle(10)=225 : angle(11)=240 : angle(12)=270 : angle(13)=300: angle(14)=315 : angle(15)=330 : angle(16)=360

'pour réduire les boucles aspects en fonction de la division par 15 des longitudes (1er chiffre : 1er aspect, 2ème chiffre : dernier aspect à considérer)
arc(0,0)=0 : arc(0,1)=0 : arc(1,0)=1 : arc(1,1)=1 : arc(2,0)=1 : arc(2,1)=2 : arc(3,0)=2 : arc(3,1)=3 : arc(4,0)=3 : arc(4,1)=3
arc(5,0)=4 : arc(5,1)=4 : arc(6,0)=4 : arc(6,1)=4 : arc(7,0)=5 : arc(7,1)=5 : arc(8,0)=5 : arc(8,1)=6 : arc(9,0)=6 : arc(9,1)=7
arc(10,0)=7 : arc(10,1)=7 : arc(11,0)=8 : arc(11,1)=8 : arc(12,0)=8 : arc(12,1)=8 : arc(13,0)=8 : arc(13,1)=9 : arc(14,0)=9 : arc(14,1)=10
arc(15,0)=10 : arc(15,1)=11 : arc(16,0)=11 : arc(16,1)=11 : arc(17,0)=12 : arc(17,1)=12 : arc(18,0)=12 : arc(18,1)=12 : arc(19,0)=13 : arc(19,1)=13
arc(20,0)=13 : arc(20,1)=14 : arc(21,0)=14 : arc(21,1)=15 : arc(22,0)=15 : arc(22,1)=15 : arc(23,0)=16 : arc(23,1)=16 '16=0 permet rebouclage
arc(24,0)=0 : arc(24,1)=0 'utile si gap=360 juste (360/15=24) ?
			
bleu= RGB(0, 0, 255) : rouge = RGB(255, 0, 0) : mauve = RGB(153, 51, 255) : violet=RGB(255,51,255)
couleur(0)= bleu : couleur(1)= bleu : couleur(2)= rouge : couleur(3)= bleu
couleur(4)= rouge : couleur(5)= bleu : couleur(6)= violet : couleur(7)= mauve
couleur(8)= rouge : couleur(9)= mauve : couleur(10)= violet : couleur(11)= bleu
couleur(12)= rouge : couleur(13)= bleu : couleur(14)= rouge : couleur(15)= bleu

signe(0)="Bélier" :  signe(1)="Taureau" : signe(2)="Gémeaux" : signe(3)="Cancer" : signe(4)="Lion" : signe(5)="Vierge" : 
signe(6)="Balance" :  signe(7)="Scorpion" : signe(8)="Sagittaire" : signe(9)="Capricorne" : signe(10)="Verseau" : signe(11)="Poissons"

orbe(0)=15 : orbe(1)=2 : orbe(2)=2 : orbe(3)=4:
orbe(4)=6 : orbe(5)=8 : orbe(6)=2 : orbe(7)=1/2:: orbe(8)=10
orbe(9)=1/2 : orbe(10)=2 : orbe(11)=8
orbe(12)=6 : orbe(13)=4 : orbe(14)=2 : orbe(15)=2

'orbe_transit(0)=2+1/2 : orbe_transit(1)=2 : orbe_transit(2)=2 : orbe_transit(3)=2:
'orbe_transit(4)=2 : orbe_transit(5)=2 : orbe_transit(6)=2 : orbe_transit(7)=1/2:: orbe_transit(8)=2
'orbe_transit(9)=1/2 : orbe_transit(10)=2 : orbe_transit(11)=2
'orbe_transit(12)=2 : orbe_transit(13)=2 : orbe_transit(14)=2 : orbe_transit(15)=2

orbe_transit(0)=1/10 : orbe_transit(1)=1/10 : orbe_transit(2)=1/10 : orbe_transit(3)=1/10
orbe_transit(4)=1/10 : orbe_transit(5)=1/10 : orbe_transit(6)=1/10 : orbe_transit(7)=1/10: orbe_transit(8)=1/10
orbe_transit(9)=1/10 : orbe_transit(10)=1/10 : orbe_transit(11)=1/10
orbe_transit(12)=1/10 : orbe_transit(13)=1/10 : orbe_transit(14)=1/10 : orbe_transit(15)=1/10

matrice (0,4)="Maîtrise": matrice (0,6)="Chute+" : matrice (0,0)="Exaltation" : matrice (0,10)="Exil" 'soleil,signe
matrice (1,3)="Maîtrise": matrice (1,7)="Chute+" : matrice (1,1)="Exaltation" : matrice (1,9)="Exil" 'lune
matrice (2,2)="Maîtrise et Chute": matrice (2,5)="Maîtrise et Exaltation" :  matrice (2,11)="Exil et Chute" 'mercure
matrice (2,8)="Exil et Exaltation"
matrice (3,1)="Maîtrise": matrice (3,6)="Maîtrise" : matrice (3,5)="Chute" :  matrice (3,8)="Chute" 'venus
matrice (3,2)="Exaltation" : matrice (3,11)="Exaltation" : matrice (3,0)="Exil" : matrice (3,7)="Exil"
matrice (4,0)="Maîtrise": matrice (4,7)="Maîtrise" : matrice (4,3)="Chute-" :  matrice (4,10)="Chute+" 'mars
matrice (4,4)="Exaltation" : matrice (4,9)="Exaltation" : matrice (4,1)="Exil" : matrice (4,6)="Exil"
matrice (5,8)="Maîtrise": matrice (5,11)="Maîtrise" : matrice (5,4)="Chute-" :  matrice (5,9)="Chute+" 'jupiter
matrice (5,3)="Exaltation" : matrice (5,10)="Exaltation" : matrice (5,2)="Exil" : matrice (5,5)="Exil"
matrice (6,9)="Maîtrise": matrice (6,10)="Maîtrise" : matrice (6,0)="Chute-" :  matrice (6,1)="Chute-" 'saturne
matrice (6,6)="Exaltation" : matrice (6,7)="Exaltation" : matrice (6,3)="Exil" : matrice (6,4)="Exil"
matrice (7,9)="Maîtrise": matrice (7,10)="Maîtrise":  matrice (7,3)="Exil" : matrice (7,4)="Exil" 'uranus
matrice (8,8)="Maîtrise": matrice (8,11)="Maîtrise" : matrice (8,2)="Exil" : matrice (0,5)="Exil" 'neptune
matrice (9,0)="Maîtrise": matrice (9,7)="Maîtrise" : matrice (9,1)="Exil" : matrice (9,6)="Exil" 'pluton

'Soleil, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 0, 0, 0)= 20.6333	: orbe_theme( 0, 0, 1)= 0.95	: orbe_theme( 0, 0, 2)= 4.15	: orbe_theme( 0, 0, 3)= 6.2	: orbe_theme( 0, 0, 4)= 11.85	: orbe_theme( 0, 0, 5)= 14.0166	: orbe_theme( 0, 0, 6)= 5.25	: orbe_theme( 0, 0, 7)= 4.6833
orbe_theme( 0, 0, 8)= 20.6333	 : orbe_theme( 0, 0, 9)= 4.6833	 : orbe_theme( 0, 0, 10)= 5.25	 : orbe_theme( 0, 0, 11)= 14.0166	 : orbe_theme( 0, 0, 12)= 11.85	 : orbe_theme( 0, 0, 13)= 6.2	 : orbe_theme( 0, 0, 14)= 4.15	 : orbe_theme( 0, 0, 15)= 0.95
orbe_theme( 0, 1, 0)= 20.6166	: orbe_theme( 0, 1, 1)= 0.95	: orbe_theme( 0, 1, 2)= 4.15	: orbe_theme( 0, 1, 3)= 6.2	: orbe_theme( 0, 1, 4)= 11.8333	: orbe_theme( 0, 1, 5)= 13.9833	: orbe_theme( 0, 1, 6)= 5.25	: orbe_theme( 0, 1, 7)= 4.6833
orbe_theme( 0, 1, 8)= 20.6166	 : orbe_theme( 0, 1, 9)= 4.6833	 : orbe_theme( 0, 1, 10)= 5.25	 : orbe_theme( 0, 1, 11)= 13.9833	 : orbe_theme( 0, 1, 12)= 11.8333	 : orbe_theme( 0, 1, 13)= 6.2	 : orbe_theme( 0, 1, 14)= 4.15	 : orbe_theme( 0, 1, 15)= 0.95
orbe_theme( 0, 2, 0)= 12.9666	: orbe_theme( 0, 2, 1)= 0.6	: orbe_theme( 0, 2, 2)= 2.6166	: orbe_theme( 0, 2, 3)= 3.9	: orbe_theme( 0, 2, 4)= 7.45	: orbe_theme( 0, 2, 5)= 8.8	: orbe_theme( 0, 2, 6)= 3.3	: orbe_theme( 0, 2, 7)= 3.5833
orbe_theme( 0, 2, 8)= 12.9666	 : orbe_theme( 0, 2, 9)= 3.5833	 : orbe_theme( 0, 2, 10)= 3.3	 : orbe_theme( 0, 2, 11)= 8.8	 : orbe_theme( 0, 2, 12)= 7.45	 : orbe_theme( 0, 2, 13)= 3.9	 : orbe_theme( 0, 2, 14)= 2.6166	 : orbe_theme( 0, 2, 15)= 0.6
orbe_theme( 0, 3, 0)= 16.6833	: orbe_theme( 0, 3, 1)= 0.7666	: orbe_theme( 0, 3, 2)= 3.35	: orbe_theme( 0, 3, 3)= 5.0166	: orbe_theme( 0, 3, 4)= 9.5666	: orbe_theme( 0, 3, 5)= 11.3333	: orbe_theme( 0, 3, 6)= 4.25	: orbe_theme( 0, 3, 7)= 4.6166
orbe_theme( 0, 3, 8)= 16.6833	 : orbe_theme( 0, 3, 9)= 4.6166	 : orbe_theme( 0, 3, 10)= 4.25	 : orbe_theme( 0, 3, 11)= 11.3333	 : orbe_theme( 0, 3, 12)= 9.5666	 : orbe_theme( 0, 3, 13)= 5.0166	 : orbe_theme( 0, 3, 14)= 3.35	 : orbe_theme( 0, 3, 15)= 0.7666
orbe_theme( 0, 4, 0)= 15.7333	: orbe_theme( 0, 4, 1)= 0.7333	: orbe_theme( 0, 4, 2)= 3.1666	: orbe_theme( 0, 4, 3)= 4.7333	: orbe_theme( 0, 4, 4)= 9.0333	: orbe_theme( 0, 4, 5)= 10.6833	: orbe_theme( 0, 4, 6)= 4	: orbe_theme( 0, 4, 7)= 4.35
orbe_theme( 0, 4, 8)= 15.7333	 : orbe_theme( 0, 4, 9)= 4.35	 : orbe_theme( 0, 4, 10)= 4	 : orbe_theme( 0, 4, 11)= 10.6833	 : orbe_theme( 0, 4, 12)= 9.0333	 : orbe_theme( 0, 4, 13)= 4.7333	 : orbe_theme( 0, 4, 14)= 3.1666	 : orbe_theme( 0, 4, 15)= 0.7333
orbe_theme( 0, 5, 0)= 16.7666	: orbe_theme( 0, 5, 1)= 0.7666	: orbe_theme( 0, 5, 2)= 3.3833	: orbe_theme( 0, 5, 3)= 5.05	: orbe_theme( 0, 5, 4)= 9.6333	: orbe_theme( 0, 5, 5)= 11.3833	: orbe_theme( 0, 5, 6)= 4.2666	: orbe_theme( 0, 5, 7)= 4.6333
orbe_theme( 0, 5, 8)= 16.7666	 : orbe_theme( 0, 5, 9)= 4.6333	 : orbe_theme( 0, 5, 10)= 4.2666	 : orbe_theme( 0, 5, 11)= 11.3833	 : orbe_theme( 0, 5, 12)= 9.6333	 : orbe_theme( 0, 5, 13)= 5.05	 : orbe_theme( 0, 5, 14)= 3.3833	 : orbe_theme( 0, 5, 15)= 0.7666
orbe_theme( 0, 6, 0)= 15.7833	: orbe_theme( 0, 6, 1)= 0.7333	: orbe_theme( 0, 6, 2)= 3.1666	: orbe_theme( 0, 6, 3)= 4.75	: orbe_theme( 0, 6, 4)= 9.05	: orbe_theme( 0, 6, 5)= 10.7166	: orbe_theme( 0, 6, 6)= 4.0166	: orbe_theme( 0, 6, 7)= 4.3666
orbe_theme( 0, 6, 8)= 15.7833	 : orbe_theme( 0, 6, 9)= 4.3666	 : orbe_theme( 0, 6, 10)= 4.0166	 : orbe_theme( 0, 6, 11)= 10.7166	 : orbe_theme( 0, 6, 12)= 9.05	 : orbe_theme( 0, 6, 13)= 4.75	 : orbe_theme( 0, 6, 14)= 3.1666	 : orbe_theme( 0, 6, 15)= 0.7333
orbe_theme( 0, 7, 0)= 14.3	: orbe_theme( 0, 7, 1)= 0.6666	: orbe_theme( 0, 7, 2)= 2.8833	: orbe_theme( 0, 7, 3)= 4.3	: orbe_theme( 0, 7, 4)= 8.2	: orbe_theme( 0, 7, 5)= 9.7	: orbe_theme( 0, 7, 6)= 3.6333	: orbe_theme( 0, 7, 7)= 3.95
orbe_theme( 0, 7, 8)= 14.3	 : orbe_theme( 0, 7, 9)= 3.95	 : orbe_theme( 0, 7, 10)= 3.6333	 : orbe_theme( 0, 7, 11)= 9.7	 : orbe_theme( 0, 7, 12)= 8.2	 : orbe_theme( 0, 7, 13)= 4.3	 : orbe_theme( 0, 7, 14)= 2.8833	 : orbe_theme( 0, 7, 15)= 0.6666
orbe_theme( 0, 8, 0)= 13.7666	: orbe_theme( 0, 8, 1)= 0.6333	: orbe_theme( 0, 8, 2)= 2.7666	: orbe_theme( 0, 8, 3)= 4.1333	: orbe_theme( 0, 8, 4)= 7.9	: orbe_theme( 0, 8, 5)= 9.35	: orbe_theme( 0, 8, 6)= 3.5	: orbe_theme( 0, 8, 7)= 3.8
orbe_theme( 0, 8, 8)= 13.7666	 : orbe_theme( 0, 8, 9)= 3.8	 : orbe_theme( 0, 8, 10)= 3.5	 : orbe_theme( 0, 8, 11)= 9.35	 : orbe_theme( 0, 8, 12)= 7.9	 : orbe_theme( 0, 8, 13)= 4.1333	 : orbe_theme( 0, 8, 14)= 2.7666	 : orbe_theme( 0, 8, 15)= 0.6333
orbe_theme( 0, 9, 0)= 13.6666	: orbe_theme( 0, 9, 1)= 0.6333	: orbe_theme( 0, 9, 2)= 2.75	: orbe_theme( 0, 9, 3)= 4.1	: orbe_theme( 0, 9, 4)= 7.8333	: orbe_theme( 0, 9, 5)= 9.2666	: orbe_theme( 0, 9, 6)= 3.4833	: orbe_theme( 0, 9, 7)= 3.7833
orbe_theme( 0, 9, 8)= 13.6666	 : orbe_theme( 0, 9, 9)= 3.7833	 : orbe_theme( 0, 9, 10)= 3.4833	 : orbe_theme( 0, 9, 11)= 9.2666	 : orbe_theme( 0, 9, 12)= 7.8333	 : orbe_theme( 0, 9, 13)= 4.1	 : orbe_theme( 0, 9, 14)= 2.75	 : orbe_theme( 0, 9, 15)= 0.6333
orbe_theme( 0, 10, 0)= 10.3166	: orbe_theme( 0, 10, 1)= 0.4666	: orbe_theme( 0, 10, 2)= 2.0666	: orbe_theme( 0, 10, 3)= 3.1	: orbe_theme( 0, 10, 4)= 5.9166	: orbe_theme( 0, 10, 5)= 7	: orbe_theme( 0, 10, 6)= 2.6166	: orbe_theme( 0, 10, 7)= 2.85
orbe_theme( 0, 10, 8)= 10.3166	 : orbe_theme( 0, 10, 9)= 2.85	 : orbe_theme( 0, 10, 10)= 2.6166	 : orbe_theme( 0, 10, 11)= 7	 : orbe_theme( 0, 10, 12)= 5.9166	 : orbe_theme( 0, 10, 13)= 3.1	 : orbe_theme( 0, 10, 14)= 2.0666	 : orbe_theme( 0, 10, 15)= 0.4666
orbe_theme( 0, 11, 0)= 10.3166	: orbe_theme( 0, 11, 1)= 0.4666	: orbe_theme( 0, 11, 2)= 2.0666	: orbe_theme( 0, 11, 3)= 3.1	: orbe_theme( 0, 11, 4)= 5.9166	: orbe_theme( 0, 11, 5)= 7	: orbe_theme( 0, 11, 6)= 2.6166	: orbe_theme( 0, 11, 7)= 2.85
orbe_theme( 0, 11, 8)= 10.3166	 : orbe_theme( 0, 11, 9)= 2.85	 : orbe_theme( 0, 11, 10)= 2.6166	 : orbe_theme( 0, 11, 11)= 7	 : orbe_theme( 0, 11, 12)= 5.9166	 : orbe_theme( 0, 11, 13)= 3.1	 : orbe_theme( 0, 11, 14)= 2.0666	 : orbe_theme( 0, 11, 15)= 0.4666


'Lune, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 1, 0, 0)= 20.6166	: orbe_theme( 1, 0, 1)= 0.95	: orbe_theme( 1, 0, 2)= 4.15	: orbe_theme( 1, 0, 3)= 6.2	: orbe_theme( 1, 0, 4)= 11.8333	: orbe_theme( 1, 0, 5)= 13.9833	: orbe_theme( 1, 0, 6)= 5.25	: orbe_theme( 1, 0, 7)= 4.6833
orbe_theme( 1, 0, 8)= 20.6166	 : orbe_theme( 1, 0, 9)= 4.6833	 : orbe_theme( 1, 0, 10)= 5.25	 : orbe_theme( 1, 0, 11)= 13.9833	 : orbe_theme( 1, 0, 12)= 11.8333	 : orbe_theme( 1, 0, 13)= 6.2	 : orbe_theme( 1, 0, 14)= 4.15	 : orbe_theme( 1, 0, 15)= 0.95
orbe_theme( 1, 1, 0)= 20.5833	: orbe_theme( 1, 1, 1)= 0.95	: orbe_theme( 1, 1, 2)= 4.15	: orbe_theme( 1, 1, 3)= 6.1833	: orbe_theme( 1, 1, 4)= 11.8166	: orbe_theme( 1, 1, 5)= 13.9666	: orbe_theme( 1, 1, 6)= 5.25	: orbe_theme( 1, 1, 7)= 4.6666
orbe_theme( 1, 1, 8)= 20.5833	 : orbe_theme( 1, 1, 9)= 4.6666	 : orbe_theme( 1, 1, 10)= 5.25	 : orbe_theme( 1, 1, 11)= 13.9666	 : orbe_theme( 1, 1, 12)= 11.8166	 : orbe_theme( 1, 1, 13)= 6.1833	 : orbe_theme( 1, 1, 14)= 4.15	 : orbe_theme( 1, 1, 15)= 0.95
orbe_theme( 1, 2, 0)= 12.95	: orbe_theme( 1, 2, 1)= 0.6	: orbe_theme( 1, 2, 2)= 2.6	: orbe_theme( 1, 2, 3)= 3.8833	: orbe_theme( 1, 2, 4)= 7.4333	: orbe_theme( 1, 2, 5)= 8.7833	: orbe_theme( 1, 2, 6)= 3.3	: orbe_theme( 1, 2, 7)= 3.5833
orbe_theme( 1, 2, 8)= 12.95	 : orbe_theme( 1, 2, 9)= 3.5833	 : orbe_theme( 1, 2, 10)= 3.3	 : orbe_theme( 1, 2, 11)= 8.7833	 : orbe_theme( 1, 2, 12)= 7.4333	 : orbe_theme( 1, 2, 13)= 3.8833	 : orbe_theme( 1, 2, 14)= 2.6	 : orbe_theme( 1, 2, 15)= 0.6
orbe_theme( 1, 3, 0)= 16.65	: orbe_theme( 1, 3, 1)= 0.7666	: orbe_theme( 1, 3, 2)= 3.35	: orbe_theme( 1, 3, 3)= 5.0166	: orbe_theme( 1, 3, 4)= 9.55	: orbe_theme( 1, 3, 5)= 11.3	: orbe_theme( 1, 3, 6)= 4.2333	: orbe_theme( 1, 3, 7)= 4.6
orbe_theme( 1, 3, 8)= 16.65	 : orbe_theme( 1, 3, 9)= 4.6	 : orbe_theme( 1, 3, 10)= 4.2333	 : orbe_theme( 1, 3, 11)= 11.3	 : orbe_theme( 1, 3, 12)= 9.55	 : orbe_theme( 1, 3, 13)= 5.0166	 : orbe_theme( 1, 3, 14)= 3.35	 : orbe_theme( 1, 3, 15)= 0.7666
orbe_theme( 1, 4, 0)= 15.7	: orbe_theme( 1, 4, 1)= 0.7166	: orbe_theme( 1, 4, 2)= 3.1666	: orbe_theme( 1, 4, 3)= 4.7166	: orbe_theme( 1, 4, 4)= 9.0166	: orbe_theme( 1, 4, 5)= 10.6666	: orbe_theme( 1, 4, 6)= 4	: orbe_theme( 1, 4, 7)= 4.35
orbe_theme( 1, 4, 8)= 15.7	 : orbe_theme( 1, 4, 9)= 4.35	 : orbe_theme( 1, 4, 10)= 4	 : orbe_theme( 1, 4, 11)= 10.6666	 : orbe_theme( 1, 4, 12)= 9.0166	 : orbe_theme( 1, 4, 13)= 4.7166	 : orbe_theme( 1, 4, 14)= 3.1666	 : orbe_theme( 1, 4, 15)= 0.7166
orbe_theme( 1, 5, 0)= 16.75	: orbe_theme( 1, 5, 1)= 0.7666	: orbe_theme( 1, 5, 2)= 3.3666	: orbe_theme( 1, 5, 3)= 5.0333	: orbe_theme( 1, 5, 4)= 9.6166	: orbe_theme( 1, 5, 5)= 11.3666	: orbe_theme( 1, 5, 6)= 4.2666	: orbe_theme( 1, 5, 7)= 4.6333
orbe_theme( 1, 5, 8)= 16.75	 : orbe_theme( 1, 5, 9)= 4.6333	 : orbe_theme( 1, 5, 10)= 4.2666	 : orbe_theme( 1, 5, 11)= 11.3666	 : orbe_theme( 1, 5, 12)= 9.6166	 : orbe_theme( 1, 5, 13)= 5.0333	 : orbe_theme( 1, 5, 14)= 3.3666	 : orbe_theme( 1, 5, 15)= 0.7666
orbe_theme( 1, 6, 0)= 15.75	: orbe_theme( 1, 6, 1)= 0.7333	: orbe_theme( 1, 6, 2)= 3.1666	: orbe_theme( 1, 6, 3)= 4.7333	: orbe_theme( 1, 6, 4)= 9.0333	: orbe_theme( 1, 6, 5)= 10.6833	: orbe_theme( 1, 6, 6)= 4.0166	: orbe_theme( 1, 6, 7)= 4.35
orbe_theme( 1, 6, 8)= 15.75	 : orbe_theme( 1, 6, 9)= 4.35	 : orbe_theme( 1, 6, 10)= 4.0166	 : orbe_theme( 1, 6, 11)= 10.6833	 : orbe_theme( 1, 6, 12)= 9.0333	 : orbe_theme( 1, 6, 13)= 4.7333	 : orbe_theme( 1, 6, 14)= 3.1666	 : orbe_theme( 1, 6, 15)= 0.7333
orbe_theme( 1, 7, 0)= 14.2666	: orbe_theme( 1, 7, 1)= 0.65	: orbe_theme( 1, 7, 2)= 2.8666	: orbe_theme( 1, 7, 3)= 4.2833	: orbe_theme( 1, 7, 4)= 8.1833	: orbe_theme( 1, 7, 5)= 9.6833	: orbe_theme( 1, 7, 6)= 3.6333	: orbe_theme( 1, 7, 7)= 3.95
orbe_theme( 1, 7, 8)= 14.2666	 : orbe_theme( 1, 7, 9)= 3.95	 : orbe_theme( 1, 7, 10)= 3.6333	 : orbe_theme( 1, 7, 11)= 9.6833	 : orbe_theme( 1, 7, 12)= 8.1833	 : orbe_theme( 1, 7, 13)= 4.2833	 : orbe_theme( 1, 7, 14)= 2.8666	 : orbe_theme( 1, 7, 15)= 0.65
orbe_theme( 1, 8, 0)= 13.75	: orbe_theme( 1, 8, 1)= 0.6333	: orbe_theme( 1, 8, 2)= 2.7666	: orbe_theme( 1, 8, 3)= 4.1333	: orbe_theme( 1, 8, 4)= 7.8833	: orbe_theme( 1, 8, 5)= 9.3333	: orbe_theme( 1, 8, 6)= 3.5	: orbe_theme( 1, 8, 7)= 3.8
orbe_theme( 1, 8, 8)= 13.75	 : orbe_theme( 1, 8, 9)= 3.8	 : orbe_theme( 1, 8, 10)= 3.5	 : orbe_theme( 1, 8, 11)= 9.3333	 : orbe_theme( 1, 8, 12)= 7.8833	 : orbe_theme( 1, 8, 13)= 4.1333	 : orbe_theme( 1, 8, 14)= 2.7666	 : orbe_theme( 1, 8, 15)= 0.6333
orbe_theme( 1, 9, 0)= 13.6333	: orbe_theme( 1, 9, 1)= 0.6333	: orbe_theme( 1, 9, 2)= 2.75	: orbe_theme( 1, 9, 3)= 4.1	: orbe_theme( 1, 9, 4)= 7.8166	: orbe_theme( 1, 9, 5)= 9.25	: orbe_theme( 1, 9, 6)= 3.4666	: orbe_theme( 1, 9, 7)= 3.7666
orbe_theme( 1, 9, 8)= 13.6333	 : orbe_theme( 1, 9, 9)= 3.7666	 : orbe_theme( 1, 9, 10)= 3.4666	 : orbe_theme( 1, 9, 11)= 9.25	 : orbe_theme( 1, 9, 12)= 7.8166	 : orbe_theme( 1, 9, 13)= 4.1	 : orbe_theme( 1, 9, 14)= 2.75	 : orbe_theme( 1, 9, 15)= 0.6333
orbe_theme( 1, 10, 0)= 10.2833	: orbe_theme( 1, 10, 1)= 0.4666	: orbe_theme( 1, 10, 2)= 2.0666	: orbe_theme( 1, 10, 3)= 3.0833	: orbe_theme( 1, 10, 4)= 5.9	: orbe_theme( 1, 10, 5)= 6.9833	: orbe_theme( 1, 10, 6)= 2.6166	: orbe_theme( 1, 10, 7)= 2.85
orbe_theme( 1, 10, 8)= 10.2833	 : orbe_theme( 1, 10, 9)= 2.85	 : orbe_theme( 1, 10, 10)= 2.6166	 : orbe_theme( 1, 10, 11)= 6.9833	 : orbe_theme( 1, 10, 12)= 5.9	 : orbe_theme( 1, 10, 13)= 3.0833	 : orbe_theme( 1, 10, 14)= 2.0666	 : orbe_theme( 1, 10, 15)= 0.4666
orbe_theme( 1, 11, 0)= 10.2833	: orbe_theme( 1, 11, 1)= 0.4666	: orbe_theme( 1, 11, 2)= 2.0666	: orbe_theme( 1, 11, 3)= 3.0833	: orbe_theme( 1, 11, 4)= 5.9	: orbe_theme( 1, 11, 5)= 6.9833	: orbe_theme( 1, 11, 6)= 2.6166	: orbe_theme( 1, 11, 7)= 2.85
orbe_theme( 1, 11, 8)= 10.2833	 : orbe_theme( 1, 11, 9)= 2.85	 : orbe_theme( 1, 11, 10)= 2.6166	 : orbe_theme( 1, 11, 11)= 6.9833	 : orbe_theme( 1, 11, 12)= 5.9	 : orbe_theme( 1, 11, 13)= 3.0833	 : orbe_theme( 1, 11, 14)= 2.0666	 : orbe_theme( 1, 11, 15)= 0.4666


'Mercure, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 2, 0, 0)= 12.9666	: orbe_theme( 2, 0, 1)= 0.6	: orbe_theme( 2, 0, 2)= 2.6166	: orbe_theme( 2, 0, 3)= 3.9	: orbe_theme( 2, 0, 4)= 7.45	: orbe_theme( 2, 0, 5)= 8.8	: orbe_theme( 2, 0, 6)= 3.3	: orbe_theme( 2, 0, 7)= 3.5833
orbe_theme( 2, 0, 8)= 12.9666	 : orbe_theme( 2, 0, 9)= 3.5833	 : orbe_theme( 2, 0, 10)= 3.3	 : orbe_theme( 2, 0, 11)= 8.8	 : orbe_theme( 2, 0, 12)= 7.45	 : orbe_theme( 2, 0, 13)= 3.9	 : orbe_theme( 2, 0, 14)= 2.6166	 : orbe_theme( 2, 0, 15)= 0.6
orbe_theme( 2, 1, 0)= 12.95	: orbe_theme( 2, 1, 1)= 0.6	: orbe_theme( 2, 1, 2)= 2.6	: orbe_theme( 2, 1, 3)= 3.8833	: orbe_theme( 2, 1, 4)= 7.4333	: orbe_theme( 2, 1, 5)= 8.7833	: orbe_theme( 2, 1, 6)= 3.3	: orbe_theme( 2, 1, 7)= 3.5833
orbe_theme( 2, 1, 8)= 12.95	 : orbe_theme( 2, 1, 9)= 3.5833	 : orbe_theme( 2, 1, 10)= 3.3	 : orbe_theme( 2, 1, 11)= 8.7833	 : orbe_theme( 2, 1, 12)= 7.4333	 : orbe_theme( 2, 1, 13)= 3.8833	 : orbe_theme( 2, 1, 14)= 2.6	 : orbe_theme( 2, 1, 15)= 0.6
orbe_theme( 2, 2, 0)= 5.3	: orbe_theme( 2, 2, 1)= 0.2333	: orbe_theme( 2, 2, 2)= 1.0666	: orbe_theme( 2, 2, 3)= 1.6	: orbe_theme( 2, 2, 4)= 3.05	: orbe_theme( 2, 2, 5)= 3.6	: orbe_theme( 2, 2, 6)= 1.35	: orbe_theme( 2, 2, 7)= 1.4666
orbe_theme( 2, 2, 8)= 5.3	 : orbe_theme( 2, 2, 9)= 1.4666	 : orbe_theme( 2, 2, 10)= 1.35	 : orbe_theme( 2, 2, 11)= 3.6	 : orbe_theme( 2, 2, 12)= 3.05	 : orbe_theme( 2, 2, 13)= 1.6	 : orbe_theme( 2, 2, 14)= 1.0666	 : orbe_theme( 2, 2, 15)= 0.2333
orbe_theme( 2, 3, 0)= 9.0166	: orbe_theme( 2, 3, 1)= 0.4166	: orbe_theme( 2, 3, 2)= 1.8166	: orbe_theme( 2, 3, 3)= 2.7166	: orbe_theme( 2, 3, 4)= 5.1666	: orbe_theme( 2, 3, 5)= 6.1166	: orbe_theme( 2, 3, 6)= 2.3	: orbe_theme( 2, 3, 7)= 2.4833
orbe_theme( 2, 3, 8)= 9.0166	 : orbe_theme( 2, 3, 9)= 2.4833	 : orbe_theme( 2, 3, 10)= 2.3	 : orbe_theme( 2, 3, 11)= 6.1166	 : orbe_theme( 2, 3, 12)= 5.1666	 : orbe_theme( 2, 3, 13)= 2.7166	 : orbe_theme( 2, 3, 14)= 1.8166	 : orbe_theme( 2, 3, 15)= 0.4166
orbe_theme( 2, 4, 0)= 8.0666	: orbe_theme( 2, 4, 1)= 0.3666	: orbe_theme( 2, 4, 2)= 1.6166	: orbe_theme( 2, 4, 3)= 2.4166	: orbe_theme( 2, 4, 4)= 4.6333	: orbe_theme( 2, 4, 5)= 5.4833	: orbe_theme( 2, 4, 6)= 2.05	: orbe_theme( 2, 4, 7)= 2.2333
orbe_theme( 2, 4, 8)= 8.0666	 : orbe_theme( 2, 4, 9)= 2.2333	 : orbe_theme( 2, 4, 10)= 2.05	 : orbe_theme( 2, 4, 11)= 5.4833	 : orbe_theme( 2, 4, 12)= 4.6333	 : orbe_theme( 2, 4, 13)= 2.4166	 : orbe_theme( 2, 4, 14)= 1.6166	 : orbe_theme( 2, 4, 15)= 0.3666
orbe_theme( 2, 5, 0)= 9.1166	: orbe_theme( 2, 5, 1)= 0.4166	: orbe_theme( 2, 5, 2)= 1.8333	: orbe_theme( 2, 5, 3)= 2.7333	: orbe_theme( 2, 5, 4)= 5.2333	: orbe_theme( 2, 5, 5)= 6.1833	: orbe_theme( 2, 5, 6)= 2.3166	: orbe_theme( 2, 5, 7)= 2.5166
orbe_theme( 2, 5, 8)= 9.1166	 : orbe_theme( 2, 5, 9)= 2.5166	 : orbe_theme( 2, 5, 10)= 2.3166	 : orbe_theme( 2, 5, 11)= 6.1833	 : orbe_theme( 2, 5, 12)= 5.2333	 : orbe_theme( 2, 5, 13)= 2.7333	 : orbe_theme( 2, 5, 14)= 1.8333	 : orbe_theme( 2, 5, 15)= 0.4166
orbe_theme( 2, 6, 0)= 8.1166	: orbe_theme( 2, 6, 1)= 0.3666	: orbe_theme( 2, 6, 2)= 1.6333	: orbe_theme( 2, 6, 3)= 2.4333	: orbe_theme( 2, 6, 4)= 4.65	: orbe_theme( 2, 6, 5)= 5.5	: orbe_theme( 2, 6, 6)= 2.0666	: orbe_theme( 2, 6, 7)= 2.2333
orbe_theme( 2, 6, 8)= 8.1166	 : orbe_theme( 2, 6, 9)= 2.2333	 : orbe_theme( 2, 6, 10)= 2.0666	 : orbe_theme( 2, 6, 11)= 5.5	 : orbe_theme( 2, 6, 12)= 4.65	 : orbe_theme( 2, 6, 13)= 2.4333	 : orbe_theme( 2, 6, 14)= 1.6333	 : orbe_theme( 2, 6, 15)= 0.3666
orbe_theme( 2, 7, 0)= 6.6333	: orbe_theme( 2, 7, 1)= 0.3	: orbe_theme( 2, 7, 2)= 1.3333	: orbe_theme( 2, 7, 3)= 2	: orbe_theme( 2, 7, 4)= 3.8	: orbe_theme( 2, 7, 5)= 4.5	: orbe_theme( 2, 7, 6)= 1.6833	: orbe_theme( 2, 7, 7)= 1.8333
orbe_theme( 2, 7, 8)= 6.6333	 : orbe_theme( 2, 7, 9)= 1.8333	 : orbe_theme( 2, 7, 10)= 1.6833	 : orbe_theme( 2, 7, 11)= 4.5	 : orbe_theme( 2, 7, 12)= 3.8	 : orbe_theme( 2, 7, 13)= 2	 : orbe_theme( 2, 7, 14)= 1.3333	 : orbe_theme( 2, 7, 15)= 0.3
orbe_theme( 2, 8, 0)= 6.1333	: orbe_theme( 2, 8, 1)= 0.2833	: orbe_theme( 2, 8, 2)= 1.2166	: orbe_theme( 2, 8, 3)= 1.8333	: orbe_theme( 2, 8, 4)= 3.5	: orbe_theme( 2, 8, 5)= 4.15	: orbe_theme( 2, 8, 6)= 1.55	: orbe_theme( 2, 8, 7)= 1.6833
orbe_theme( 2, 8, 8)= 6.1	 : orbe_theme( 2, 8, 9)= 1.6833	 : orbe_theme( 2, 8, 10)= 1.55	 : orbe_theme( 2, 8, 11)= 4.15	 : orbe_theme( 2, 8, 12)= 3.5	 : orbe_theme( 2, 8, 13)= 1.8333	 : orbe_theme( 2, 8, 14)= 1.2166	 : orbe_theme( 2, 8, 15)= 0.2833
orbe_theme( 2, 9, 0)= 6	: orbe_theme( 2, 9, 1)= 0.2666	: orbe_theme( 2, 9, 2)= 1.2	: orbe_theme( 2, 9, 3)= 1.8	: orbe_theme( 2, 9, 4)= 3.4333	: orbe_theme( 2, 9, 5)= 4.0666	: orbe_theme( 2, 9, 6)= 1.5166	: orbe_theme( 2, 9, 7)= 1.65
orbe_theme( 2, 9, 8)= 6	 : orbe_theme( 2, 9, 9)= 1.65	 : orbe_theme( 2, 9, 10)= 1.5166	 : orbe_theme( 2, 9, 11)= 4.0666	 : orbe_theme( 2, 9, 12)= 3.4333	 : orbe_theme( 2, 9, 13)= 1.8	 : orbe_theme( 2, 9, 14)= 1.2	 : orbe_theme( 2, 9, 15)= 0.2666
orbe_theme( 2, 10, 0)= 2.65	: orbe_theme( 2, 10, 1)= 0.1166	: orbe_theme( 2, 10, 2)= 0.5333	: orbe_theme( 2, 10, 3)= 0.8	: orbe_theme( 2, 10, 4)= 1.5166	: orbe_theme( 2, 10, 5)= 1.8	: orbe_theme( 2, 10, 6)= 0.6666	: orbe_theme( 2, 10, 7)= 0.7333
orbe_theme( 2, 10, 8)= 2.65	 : orbe_theme( 2, 10, 9)= 0.7333	 : orbe_theme( 2, 10, 10)= 0.6666	 : orbe_theme( 2, 10, 11)= 1.8	 : orbe_theme( 2, 10, 12)= 1.5166	 : orbe_theme( 2, 10, 13)= 0.8	 : orbe_theme( 2, 10, 14)= 0.5333	 : orbe_theme( 2, 10, 15)= 0.1166
orbe_theme( 2, 11, 0)= 2.65	: orbe_theme( 2, 11, 1)= 0.1166	: orbe_theme( 2, 11, 2)= 0.5333	: orbe_theme( 2, 11, 3)= 0.8	: orbe_theme( 2, 11, 4)= 1.5166	: orbe_theme( 2, 11, 5)= 1.8	: orbe_theme( 2, 11, 6)= 0.6666	: orbe_theme( 2, 11, 7)= 0.7333
orbe_theme( 2, 11, 8)= 2.65	 : orbe_theme( 2, 11, 9)= 0.7333	 : orbe_theme( 2, 11, 10)= 0.6666	 : orbe_theme( 2, 11, 11)= 1.8	 : orbe_theme( 2, 11, 12)= 1.5166	 : orbe_theme( 2, 11, 13)= 0.8	 : orbe_theme( 2, 11, 14)= 0.5333	 : orbe_theme( 2, 11, 15)= 0.1166


'Vénus, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 3, 0, 0)= 16.6833	: orbe_theme( 3, 0, 1)= 0.7666	: orbe_theme( 3, 0, 2)= 3.35	: orbe_theme( 3, 0, 3)= 5.0166	: orbe_theme( 3, 0, 4)= 9.5666	: orbe_theme( 3, 0, 5)= 11.3333	: orbe_theme( 3, 0, 6)= 4.25	: orbe_theme( 3, 0, 7)= 4.6166
orbe_theme( 3, 0, 8)= 16.6833	 : orbe_theme( 3, 0, 9)= 4.6166	 : orbe_theme( 3, 0, 10)= 4.25	 : orbe_theme( 3, 0, 11)= 11.3333	 : orbe_theme( 3, 0, 12)= 9.5666	 : orbe_theme( 3, 0, 13)= 5.0166	 : orbe_theme( 3, 0, 14)= 3.35	 : orbe_theme( 3, 0, 15)= 0.7666
orbe_theme( 3, 1, 0)= 16.65	: orbe_theme( 3, 1, 1)= 0.7666	: orbe_theme( 3, 1, 2)= 3.35	: orbe_theme( 3, 1, 3)= 5.0166	: orbe_theme( 3, 1, 4)= 9.55	: orbe_theme( 3, 1, 5)= 11.3	: orbe_theme( 3, 1, 6)= 4.2333	: orbe_theme( 3, 1, 7)= 4.6
orbe_theme( 3, 1, 8)= 16.65	 : orbe_theme( 3, 1, 9)= 4.6	 : orbe_theme( 3, 1, 10)= 4.2333	 : orbe_theme( 3, 1, 11)= 11.3	 : orbe_theme( 3, 1, 12)= 9.55	 : orbe_theme( 3, 1, 13)= 5.0166	 : orbe_theme( 3, 1, 14)= 3.35	 : orbe_theme( 3, 1, 15)= 0.7666
orbe_theme( 3, 2, 0)= 9.0166	: orbe_theme( 3, 2, 1)= 0.4166	: orbe_theme( 3, 2, 2)= 1.8166	: orbe_theme( 3, 2, 3)= 2.7166	: orbe_theme( 3, 2, 4)= 5.1666	: orbe_theme( 3, 2, 5)= 6.1166	: orbe_theme( 3, 2, 6)= 2.3	: orbe_theme( 3, 2, 7)= 2.4833
orbe_theme( 3, 2, 8)= 9.0166	 : orbe_theme( 3, 2, 9)= 2.4833	 : orbe_theme( 3, 2, 10)= 2.3	 : orbe_theme( 3, 2, 11)= 6.1166	 : orbe_theme( 3, 2, 12)= 5.1666	 : orbe_theme( 3, 2, 13)= 2.7166	 : orbe_theme( 3, 2, 14)= 1.8166	 : orbe_theme( 3, 2, 15)= 0.4166
orbe_theme( 3, 3, 0)= 12.7333	: orbe_theme( 3, 3, 1)= 0.5833	: orbe_theme( 3, 3, 2)= 2.5666	: orbe_theme( 3, 3, 3)= 3.8333	: orbe_theme( 3, 3, 4)= 7.3	: orbe_theme( 3, 3, 5)= 8.6333	: orbe_theme( 3, 3, 6)= 3.2333	: orbe_theme( 3, 3, 7)= 3.5166
orbe_theme( 3, 3, 8)= 12.7333	 : orbe_theme( 3, 3, 9)= 3.5166	 : orbe_theme( 3, 3, 10)= 3.2333	 : orbe_theme( 3, 3, 11)= 8.6333	 : orbe_theme( 3, 3, 12)= 7.3	 : orbe_theme( 3, 3, 13)= 3.8333	 : orbe_theme( 3, 3, 14)= 2.5666	 : orbe_theme( 3, 3, 15)= 0.5833
orbe_theme( 3, 4, 0)= 11.7833	: orbe_theme( 3, 4, 1)= 0.55	: orbe_theme( 3, 4, 2)= 2.3666	: orbe_theme( 3, 4, 3)= 3.55	: orbe_theme( 3, 4, 4)= 6.7666	: orbe_theme( 3, 4, 5)= 8	: orbe_theme( 3, 4, 6)= 3	: orbe_theme( 3, 4, 7)= 3.25
orbe_theme( 3, 4, 8)= 11.7833	 : orbe_theme( 3, 4, 9)= 3.25	 : orbe_theme( 3, 4, 10)= 3	 : orbe_theme( 3, 4, 11)= 8	 : orbe_theme( 3, 4, 12)= 6.7666	 : orbe_theme( 3, 4, 13)= 3.55	 : orbe_theme( 3, 4, 14)= 2.3666	 : orbe_theme( 3, 4, 15)= 0.55
orbe_theme( 3, 5, 0)= 12.8166	: orbe_theme( 3, 5, 1)= 0.5833	: orbe_theme( 3, 5, 2)= 2.5833	: orbe_theme( 3, 5, 3)= 3.85	: orbe_theme( 3, 5, 4)= 7.35	: orbe_theme( 3, 5, 5)= 8.7	: orbe_theme( 3, 5, 6)= 3.2666	: orbe_theme( 3, 5, 7)= 3.55
orbe_theme( 3, 5, 8)= 12.8166	 : orbe_theme( 3, 5, 9)= 3.55	 : orbe_theme( 3, 5, 10)= 3.2666	 : orbe_theme( 3, 5, 11)= 8.7	 : orbe_theme( 3, 5, 12)= 7.35	 : orbe_theme( 3, 5, 13)= 3.85	 : orbe_theme( 3, 5, 14)= 2.5833	 : orbe_theme( 3, 5, 15)= 0.5833
orbe_theme( 3, 6, 0)= 11.8166	: orbe_theme( 3, 6, 1)= 0.55	: orbe_theme( 3, 6, 2)= 2.3833	: orbe_theme( 3, 6, 3)= 3.55	: orbe_theme( 3, 6, 4)= 6.7833	: orbe_theme( 3, 6, 5)= 8.0166	: orbe_theme( 3, 6, 6)= 3	: orbe_theme( 3, 6, 7)= 3.2666
orbe_theme( 3, 6, 8)= 11.8166	 : orbe_theme( 3, 6, 9)= 3.2666	 : orbe_theme( 3, 6, 10)= 3	 : orbe_theme( 3, 6, 11)= 8.0166	 : orbe_theme( 3, 6, 12)= 6.7833	 : orbe_theme( 3, 6, 13)= 3.55	 : orbe_theme( 3, 6, 14)= 2.3833	 : orbe_theme( 3, 6, 15)= 0.55
orbe_theme( 3, 7, 0)= 10.35	: orbe_theme( 3, 7, 1)= 0.4666	: orbe_theme( 3, 7, 2)= 2.0833	: orbe_theme( 3, 7, 3)= 3.1166	: orbe_theme( 3, 7, 4)= 5.9333	: orbe_theme( 3, 7, 5)= 7.0166	: orbe_theme( 3, 7, 6)= 2.6333	: orbe_theme( 3, 7, 7)= 2.8666
orbe_theme( 3, 7, 8)= 10.35	 : orbe_theme( 3, 7, 9)= 2.8666	 : orbe_theme( 3, 7, 10)= 2.6333	 : orbe_theme( 3, 7, 11)= 7.0166	 : orbe_theme( 3, 7, 12)= 5.9333	 : orbe_theme( 3, 7, 13)= 3.1166	 : orbe_theme( 3, 7, 14)= 2.0833	 : orbe_theme( 3, 7, 15)= 0.4666
orbe_theme( 3, 8, 0)= 9.8166	: orbe_theme( 3, 8, 1)= 0.45	: orbe_theme( 3, 8, 2)= 1.9666	: orbe_theme( 3, 8, 3)= 2.95	: orbe_theme( 3, 8, 4)= 5.6333	: orbe_theme( 3, 8, 5)= 6.6666	: orbe_theme( 3, 8, 6)= 2.5	: orbe_theme( 3, 8, 7)= 2.7166
orbe_theme( 3, 8, 8)= 9.8166	 : orbe_theme( 3, 8, 9)= 2.7166	 : orbe_theme( 3, 8, 10)= 2.5	 : orbe_theme( 3, 8, 11)= 6.6666	 : orbe_theme( 3, 8, 12)= 5.6333	 : orbe_theme( 3, 8, 13)= 2.95	 : orbe_theme( 3, 8, 14)= 1.9666	 : orbe_theme( 3, 8, 15)= 0.45
orbe_theme( 3, 9, 0)= 9.7166	: orbe_theme( 3, 9, 1)= 0.45	: orbe_theme( 3, 9, 2)= 1.95	: orbe_theme( 3, 9, 3)= 2.9166	: orbe_theme( 3, 9, 4)= 5.5666	: orbe_theme( 3, 9, 5)= 6.5833	: orbe_theme( 3, 9, 6)= 2.4666	: orbe_theme( 3, 9, 7)= 2.6833
orbe_theme( 3, 9, 8)= 9.7166	 : orbe_theme( 3, 9, 9)= 2.6833	 : orbe_theme( 3, 9, 10)= 2.4666	 : orbe_theme( 3, 9, 11)= 6.5833	 : orbe_theme( 3, 9, 12)= 5.5666	 : orbe_theme( 3, 9, 13)= 2.9166	 : orbe_theme( 3, 9, 14)= 1.95	 : orbe_theme( 3, 9, 15)= 0.45
orbe_theme( 3, 10, 0)= 6.3666	: orbe_theme( 3, 10, 1)= 0.2833	: orbe_theme( 3, 10, 2)= 1.2833	: orbe_theme( 3, 10, 3)= 1.9166	: orbe_theme( 3, 10, 4)= 3.65	: orbe_theme( 3, 10, 5)= 4.3166	: orbe_theme( 3, 10, 6)= 1.6166	: orbe_theme( 3, 10, 7)= 1.75
orbe_theme( 3, 10, 8)= 6.3666	 : orbe_theme( 3, 10, 9)= 1.75	 : orbe_theme( 3, 10, 10)= 1.6166	 : orbe_theme( 3, 10, 11)= 4.3166	 : orbe_theme( 3, 10, 12)= 3.65	 : orbe_theme( 3, 10, 13)= 1.9166	 : orbe_theme( 3, 10, 14)= 1.2833	 : orbe_theme( 3, 10, 15)= 0.2833
orbe_theme( 3, 11, 0)= 6.3666	: orbe_theme( 3, 11, 1)= 0.2833	: orbe_theme( 3, 11, 2)= 1.2833	: orbe_theme( 3, 11, 3)= 1.9166	: orbe_theme( 3, 11, 4)= 3.65	: orbe_theme( 3, 11, 5)= 4.3166	: orbe_theme( 3, 11, 6)= 1.6166	: orbe_theme( 3, 11, 7)= 1.75
orbe_theme( 3, 11, 8)= 6.3666	 : orbe_theme( 3, 11, 9)= 1.75	 : orbe_theme( 3, 11, 10)= 1.6166	 : orbe_theme( 3, 11, 11)= 4.3166	 : orbe_theme( 3, 11, 12)= 3.65	 : orbe_theme( 3, 11, 13)= 1.9166	 : orbe_theme( 3, 11, 14)= 1.2833	 : orbe_theme( 3, 11, 15)= 0.2833


'Mars, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 4, 0, 0)= 15.7333	: orbe_theme( 4, 0, 1)= 0.7333	: orbe_theme( 4, 0, 2)= 3.1666	: orbe_theme( 4, 0, 3)= 4.7333	: orbe_theme( 4, 0, 4)= 9.0333	: orbe_theme( 4, 0, 5)= 10.6833	: orbe_theme( 4, 0, 6)= 4	: orbe_theme( 4, 0, 7)= 4.35
orbe_theme( 4, 0, 8)= 15.7333	 : orbe_theme( 4, 0, 9)= 4.35	 : orbe_theme( 4, 0, 10)= 4	 : orbe_theme( 4, 0, 11)= 10.6833	 : orbe_theme( 4, 0, 12)= 9.0333	 : orbe_theme( 4, 0, 13)= 4.7333	 : orbe_theme( 4, 0, 14)= 3.1666	 : orbe_theme( 4, 0, 15)= 0.7333
orbe_theme( 4, 1, 0)= 15.7	: orbe_theme( 4, 1, 1)= 0.7166	: orbe_theme( 4, 1, 2)= 3.1666	: orbe_theme( 4, 1, 3)= 4.7166	: orbe_theme( 4, 1, 4)= 9.0166	: orbe_theme( 4, 1, 5)= 10.6666	: orbe_theme( 4, 1, 6)= 4	: orbe_theme( 4, 1, 7)= 4.35
orbe_theme( 4, 1, 8)= 15.7	 : orbe_theme( 4, 1, 9)= 4.35	 : orbe_theme( 4, 1, 10)= 4	 : orbe_theme( 4, 1, 11)= 10.6666	 : orbe_theme( 4, 1, 12)= 9.0166	 : orbe_theme( 4, 1, 13)= 4.7166	 : orbe_theme( 4, 1, 14)= 3.1666	 : orbe_theme( 4, 1, 15)= 0.7166
orbe_theme( 4, 2, 0)= 8.0666	: orbe_theme( 4, 2, 1)= 0.3666	: orbe_theme( 4, 2, 2)= 1.6166	: orbe_theme( 4, 2, 3)= 2.4166	: orbe_theme( 4, 2, 4)= 4.6333	: orbe_theme( 4, 2, 5)= 5.4833	: orbe_theme( 4, 2, 6)= 2.05	: orbe_theme( 4, 2, 7)= 2.2333
orbe_theme( 4, 2, 8)= 8.0666	 : orbe_theme( 4, 2, 9)= 2.2333	 : orbe_theme( 4, 2, 10)= 2.05	 : orbe_theme( 4, 2, 11)= 5.4833	 : orbe_theme( 4, 2, 12)= 4.6333	 : orbe_theme( 4, 2, 13)= 2.4166	 : orbe_theme( 4, 2, 14)= 1.6166	 : orbe_theme( 4, 2, 15)= 0.3666
orbe_theme( 4, 3, 0)= 11.7833	: orbe_theme( 4, 3, 1)= 0.55	: orbe_theme( 4, 3, 2)= 2.3666	: orbe_theme( 4, 3, 3)= 3.55	: orbe_theme( 4, 3, 4)= 5.75	: orbe_theme( 4, 3, 5)= 8	: orbe_theme( 4, 3, 6)= 3	: orbe_theme( 4, 3, 7)= 3.25
orbe_theme( 4, 3, 8)= 11.7833	 : orbe_theme( 4, 3, 9)= 3.25	 : orbe_theme( 4, 3, 10)= 3	 : orbe_theme( 4, 3, 11)= 8	 : orbe_theme( 4, 3, 12)= 5.75	 : orbe_theme( 4, 3, 13)= 3.55	 : orbe_theme( 4, 3, 14)= 2.3666	 : orbe_theme( 4, 3, 15)= 0.55
orbe_theme( 4, 4, 0)= 10.8333	: orbe_theme( 4, 4, 1)= 0.5	: orbe_theme( 4, 4, 2)= 2.1833	: orbe_theme( 4, 4, 3)= 3.25	: orbe_theme( 4, 4, 4)= 6.2166	: orbe_theme( 4, 4, 5)= 7.35	: orbe_theme( 4, 4, 6)= 2.75	: orbe_theme( 4, 4, 7)= 3
orbe_theme( 4, 4, 8)= 10.8333	 : orbe_theme( 4, 4, 9)= 3	 : orbe_theme( 4, 4, 10)= 2.75	 : orbe_theme( 4, 4, 11)= 7.35	 : orbe_theme( 4, 4, 12)= 6.2166	 : orbe_theme( 4, 4, 13)= 3.25	 : orbe_theme( 4, 4, 14)= 2.1833	 : orbe_theme( 4, 4, 15)= 0.5
orbe_theme( 4, 5, 0)= 11.8666	: orbe_theme( 4, 5, 1)= 0.55	: orbe_theme( 4, 5, 2)= 2.3833	: orbe_theme( 4, 5, 3)= 3.5666	: orbe_theme( 4, 5, 4)= 5.8166	: orbe_theme( 4, 5, 5)= 8.0666	: orbe_theme( 4, 5, 6)= 3.0166	: orbe_theme( 4, 5, 7)= 3.2833
orbe_theme( 4, 5, 8)= 11.8666	 : orbe_theme( 4, 5, 9)= 3.2833	 : orbe_theme( 4, 5, 10)= 3.0166	 : orbe_theme( 4, 5, 11)= 8.0666	 : orbe_theme( 4, 5, 12)= 5.8166	 : orbe_theme( 4, 5, 13)= 3.5666	 : orbe_theme( 4, 5, 14)= 2.3833	 : orbe_theme( 4, 5, 15)= 0.55
orbe_theme( 4, 6, 0)= 10.8833	: orbe_theme( 4, 6, 1)= 0.5	: orbe_theme( 4, 6, 2)= 2.1833	: orbe_theme( 4, 6, 3)= 3.2666	: orbe_theme( 4, 6, 4)= 5.2333	: orbe_theme( 4, 6, 5)= 7.3833	: orbe_theme( 4, 6, 6)= 2.75	: orbe_theme( 4, 6, 7)= 3
orbe_theme( 4, 6, 8)= 10.8833	 : orbe_theme( 4, 6, 9)= 3	 : orbe_theme( 4, 6, 10)= 2.75	 : orbe_theme( 4, 6, 11)= 7.3833	 : orbe_theme( 4, 6, 12)= 5.2333	 : orbe_theme( 4, 6, 13)= 3.2666	 : orbe_theme( 4, 6, 14)= 2.1833	 : orbe_theme( 4, 6, 15)= 0.5
orbe_theme( 4, 7, 0)= 9.4	: orbe_theme( 4, 7, 1)= 0.4166	: orbe_theme( 4, 7, 2)= 1.8833	: orbe_theme( 4, 7, 3)= 2.8166	: orbe_theme( 4, 7, 4)= 5.4	: orbe_theme( 4, 7, 5)= 5.3833	: orbe_theme( 4, 7, 6)= 2.3833	: orbe_theme( 4, 7, 7)= 2.5833
orbe_theme( 4, 7, 8)= 9.4	 : orbe_theme( 4, 7, 9)= 2.5833	 : orbe_theme( 4, 7, 10)= 2.3833	 : orbe_theme( 4, 7, 11)= 5.3833	 : orbe_theme( 4, 7, 12)= 5.4	 : orbe_theme( 4, 7, 13)= 2.8166	 : orbe_theme( 4, 7, 14)= 1.8833	 : orbe_theme( 4, 7, 15)= 0.4166
orbe_theme( 4, 8, 0)= 8.8666	: orbe_theme( 4, 8, 1)= 0.4	: orbe_theme( 4, 8, 2)= 1.7833	: orbe_theme( 4, 8, 3)= 2.6666	: orbe_theme( 4, 8, 4)= 5.0833	: orbe_theme( 4, 8, 5)= 5.0166	: orbe_theme( 4, 8, 6)= 2.25	: orbe_theme( 4, 8, 7)= 2.45
orbe_theme( 4, 8, 8)= 8.8666	 : orbe_theme( 4, 8, 9)= 2.45	 : orbe_theme( 4, 8, 10)= 2.25	 : orbe_theme( 4, 8, 11)= 5.0166	 : orbe_theme( 4, 8, 12)= 5.0833	 : orbe_theme( 4, 8, 13)= 2.6666	 : orbe_theme( 4, 8, 14)= 1.7833	 : orbe_theme( 4, 8, 15)= 0.4
orbe_theme( 4, 9, 0)= 8.75	: orbe_theme( 4, 9, 1)= 0.4	: orbe_theme( 4, 9, 2)= 1.75	: orbe_theme( 4, 9, 3)= 2.6333	: orbe_theme( 4, 9, 4)= 5.0333	: orbe_theme( 4, 9, 5)= 5.95	: orbe_theme( 4, 9, 6)= 2.2333	: orbe_theme( 4, 9, 7)= 2.4166
orbe_theme( 4, 9, 8)= 8.75	 : orbe_theme( 4, 9, 9)= 2.4166	 : orbe_theme( 4, 9, 10)= 2.2333	 : orbe_theme( 4, 9, 11)= 5.95	 : orbe_theme( 4, 9, 12)= 5.0333	 : orbe_theme( 4, 9, 13)= 2.6333	 : orbe_theme( 4, 9, 14)= 1.75	 : orbe_theme( 4, 9, 15)= 0.4
orbe_theme( 4, 10, 0)= 5.4166	: orbe_theme( 4, 10, 1)= 0.25	: orbe_theme( 4, 10, 2)= 1.0833	: orbe_theme( 4, 10, 3)= 1.6166	: orbe_theme( 4, 10, 4)= 3.0833	: orbe_theme( 4, 10, 5)= 3.6666	: orbe_theme( 4, 10, 6)= 1.3666	: orbe_theme( 4, 10, 7)= 1.5
orbe_theme( 4, 10, 8)= 5.4166	 : orbe_theme( 4, 10, 9)= 1.5	 : orbe_theme( 4, 10, 10)= 1.3666	 : orbe_theme( 4, 10, 11)= 3.6666	 : orbe_theme( 4, 10, 12)= 3.0833	 : orbe_theme( 4, 10, 13)= 1.6166	 : orbe_theme( 4, 10, 14)= 1.0833	 : orbe_theme( 4, 10, 15)= 0.25
orbe_theme( 4, 11, 0)= 5.4166	: orbe_theme( 4, 11, 1)= 0.25	: orbe_theme( 4, 11, 2)= 1.0833	: orbe_theme( 4, 11, 3)= 1.6166	: orbe_theme( 4, 11, 4)= 3.0833	: orbe_theme( 4, 11, 5)= 3.6666	: orbe_theme( 4, 11, 6)= 1.3666	: orbe_theme( 4, 11, 7)= 1.5
orbe_theme( 4, 11, 8)= 5.4166	 : orbe_theme( 4, 11, 9)= 1.5	 : orbe_theme( 4, 11, 10)= 1.3666	 : orbe_theme( 4, 11, 11)= 3.6666	 : orbe_theme( 4, 11, 12)= 3.0833	 : orbe_theme( 4, 11, 13)= 1.6166	 : orbe_theme( 4, 11, 14)= 1.0833	 : orbe_theme( 4, 11, 15)= 0.25


'Jupiter, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 5, 0, 0)= 15.75	: orbe_theme( 5, 0, 1)= 0.75	: orbe_theme( 5, 0, 2)= 3.3833	: orbe_theme( 5, 0, 3)= 5.05	: orbe_theme( 5, 0, 4)= 9.6333	: orbe_theme( 5, 0, 5)= 11.3833	: orbe_theme( 5, 0, 6)= 4.2666	: orbe_theme( 5, 0, 7)= 4.6333
orbe_theme( 5, 0, 8)= 15.75	 : orbe_theme( 5, 0, 9)= 4.6333	 : orbe_theme( 5, 0, 10)= 4.2666	 : orbe_theme( 5, 0, 11)= 11.3833	 : orbe_theme( 5, 0, 12)= 9.6333	 : orbe_theme( 5, 0, 13)= 5.05	 : orbe_theme( 5, 0, 14)= 3.3833	 : orbe_theme( 5, 0, 15)= 0.75
orbe_theme( 5, 1, 0)= 15.75	: orbe_theme( 5, 1, 1)= 0.75	: orbe_theme( 5, 1, 2)= 3.3666	: orbe_theme( 5, 1, 3)= 5.0333	: orbe_theme( 5, 1, 4)= 9.6166	: orbe_theme( 5, 1, 5)= 11.3666	: orbe_theme( 5, 1, 6)= 4.25	: orbe_theme( 5, 1, 7)= 4.6333
orbe_theme( 5, 1, 8)= 15.75	 : orbe_theme( 5, 1, 9)= 4.6333	 : orbe_theme( 5, 1, 10)= 4.25	 : orbe_theme( 5, 1, 11)= 11.3666	 : orbe_theme( 5, 1, 12)= 9.6166	 : orbe_theme( 5, 1, 13)= 5.0333	 : orbe_theme( 5, 1, 14)= 3.3666	 : orbe_theme( 5, 1, 15)= 0.75
orbe_theme( 5, 2, 0)= 9.1166	: orbe_theme( 5, 2, 1)= 0.4166	: orbe_theme( 5, 2, 2)= 1.8333	: orbe_theme( 5, 2, 3)= 2.7333	: orbe_theme( 5, 2, 4)= 5.2333	: orbe_theme( 5, 2, 5)= 5.1833	: orbe_theme( 5, 2, 6)= 2.3166	: orbe_theme( 5, 2, 7)= 2.5166
orbe_theme( 5, 2, 8)= 9.1166	 : orbe_theme( 5, 2, 9)= 2.5166	 : orbe_theme( 5, 2, 10)= 2.3166	 : orbe_theme( 5, 2, 11)= 5.1833	 : orbe_theme( 5, 2, 12)= 5.2333	 : orbe_theme( 5, 2, 13)= 2.7333	 : orbe_theme( 5, 2, 14)= 1.8333	 : orbe_theme( 5, 2, 15)= 0.4166
orbe_theme( 5, 3, 0)= 12.8166	: orbe_theme( 5, 3, 1)= 0.5833	: orbe_theme( 5, 3, 2)= 2.5833	: orbe_theme( 5, 3, 3)= 3.85	: orbe_theme( 5, 3, 4)= 7.35	: orbe_theme( 5, 3, 5)= 8.7	: orbe_theme( 5, 3, 6)= 3.25	: orbe_theme( 5, 3, 7)= 3.55
orbe_theme( 5, 3, 8)= 12.8166	 : orbe_theme( 5, 3, 9)= 3.55	 : orbe_theme( 5, 3, 10)= 3.25	 : orbe_theme( 5, 3, 11)= 8.7	 : orbe_theme( 5, 3, 12)= 7.35	 : orbe_theme( 5, 3, 13)= 3.85	 : orbe_theme( 5, 3, 14)= 2.5833	 : orbe_theme( 5, 3, 15)= 0.5833
orbe_theme( 5, 4, 0)= 11.8666	: orbe_theme( 5, 4, 1)= 0.55	: orbe_theme( 5, 4, 2)= 2.3833	: orbe_theme( 5, 4, 3)= 3.5666	: orbe_theme( 5, 4, 4)= 5.8166	: orbe_theme( 5, 4, 5)= 8.0666	: orbe_theme( 5, 4, 6)= 3.0166	: orbe_theme( 5, 4, 7)= 3.2833
orbe_theme( 5, 4, 8)= 11.8666	 : orbe_theme( 5, 4, 9)= 3.2833	 : orbe_theme( 5, 4, 10)= 3.0166	 : orbe_theme( 5, 4, 11)= 8.0666	 : orbe_theme( 5, 4, 12)= 5.8166	 : orbe_theme( 5, 4, 13)= 3.5666	 : orbe_theme( 5, 4, 14)= 2.3833	 : orbe_theme( 5, 4, 15)= 0.55
orbe_theme( 5, 5, 0)= 12.9166	: orbe_theme( 5, 5, 1)= 0.5833	: orbe_theme( 5, 5, 2)= 2.5833	: orbe_theme( 5, 5, 3)= 3.8833	: orbe_theme( 5, 5, 4)= 7.4	: orbe_theme( 5, 5, 5)= 8.75	: orbe_theme( 5, 5, 6)= 3.2833	: orbe_theme( 5, 5, 7)= 3.5666
orbe_theme( 5, 5, 8)= 12.9166	 : orbe_theme( 5, 5, 9)= 3.5666	 : orbe_theme( 5, 5, 10)= 3.2833	 : orbe_theme( 5, 5, 11)= 8.75	 : orbe_theme( 5, 5, 12)= 7.4	 : orbe_theme( 5, 5, 13)= 3.8833	 : orbe_theme( 5, 5, 14)= 2.5833	 : orbe_theme( 5, 5, 15)= 0.5833
orbe_theme( 5, 6, 0)= 11.9166	: orbe_theme( 5, 6, 1)= 0.55	: orbe_theme( 5, 6, 2)= 2.4	: orbe_theme( 5, 6, 3)= 3.5833	: orbe_theme( 5, 6, 4)= 5.8333	: orbe_theme( 5, 6, 5)= 8.0833	: orbe_theme( 5, 6, 6)= 3.0333	: orbe_theme( 5, 6, 7)= 3.3
orbe_theme( 5, 6, 8)= 11.9166	 : orbe_theme( 5, 6, 9)= 3.3	 : orbe_theme( 5, 6, 10)= 3.0333	 : orbe_theme( 5, 6, 11)= 8.0833	 : orbe_theme( 5, 6, 12)= 5.8333	 : orbe_theme( 5, 6, 13)= 3.5833	 : orbe_theme( 5, 6, 14)= 2.4	 : orbe_theme( 5, 6, 15)= 0.55
orbe_theme( 5, 7, 0)= 10.4166	: orbe_theme( 5, 7, 1)= 0.4833	: orbe_theme( 5, 7, 2)= 2.0833	: orbe_theme( 5, 7, 3)= 3.1333	: orbe_theme( 5, 7, 4)= 5.9833	: orbe_theme( 5, 7, 5)= 7.0833	: orbe_theme( 5, 7, 6)= 2.65	: orbe_theme( 5, 7, 7)= 2.8833
orbe_theme( 5, 7, 8)= 10.4166	 : orbe_theme( 5, 7, 9)= 2.8833	 : orbe_theme( 5, 7, 10)= 2.65	 : orbe_theme( 5, 7, 11)= 7.0833	 : orbe_theme( 5, 7, 12)= 5.9833	 : orbe_theme( 5, 7, 13)= 3.1333	 : orbe_theme( 5, 7, 14)= 2.0833	 : orbe_theme( 5, 7, 15)= 0.4833
orbe_theme( 5, 8, 0)= 9.9	: orbe_theme( 5, 8, 1)= 0.45	: orbe_theme( 5, 8, 2)= 1.9833	: orbe_theme( 5, 8, 3)= 2.9833	: orbe_theme( 5, 8, 4)= 5.6833	: orbe_theme( 5, 8, 5)= 5.7166	: orbe_theme( 5, 8, 6)= 2.5166	: orbe_theme( 5, 8, 7)= 2.7333
orbe_theme( 5, 8, 8)= 9.9	 : orbe_theme( 5, 8, 9)= 2.7333	 : orbe_theme( 5, 8, 10)= 2.5166	 : orbe_theme( 5, 8, 11)= 5.7166	 : orbe_theme( 5, 8, 12)= 5.6833	 : orbe_theme( 5, 8, 13)= 2.9833	 : orbe_theme( 5, 8, 14)= 1.9833	 : orbe_theme( 5, 8, 15)= 0.45
orbe_theme( 5, 9, 0)= 9.8	: orbe_theme( 5, 9, 1)= 0.45	: orbe_theme( 5, 9, 2)= 1.9666	: orbe_theme( 5, 9, 3)= 2.95	: orbe_theme( 5, 9, 4)= 5.6166	: orbe_theme( 5, 9, 5)= 5.65	: orbe_theme( 5, 9, 6)= 2.5	: orbe_theme( 5, 9, 7)= 2.7
orbe_theme( 5, 9, 8)= 9.8	 : orbe_theme( 5, 9, 9)= 2.7	 : orbe_theme( 5, 9, 10)= 2.5	 : orbe_theme( 5, 9, 11)= 5.65	 : orbe_theme( 5, 9, 12)= 5.6166	 : orbe_theme( 5, 9, 13)= 2.95	 : orbe_theme( 5, 9, 14)= 1.9666	 : orbe_theme( 5, 9, 15)= 0.45
orbe_theme( 5, 10, 0)= 5.45	: orbe_theme( 5, 10, 1)= 0.3	: orbe_theme( 5, 10, 2)= 1.3	: orbe_theme( 5, 10, 3)= 1.9166	: orbe_theme( 5, 10, 4)= 3.7	: orbe_theme( 5, 10, 5)= 4.3833	: orbe_theme( 5, 10, 6)= 1.6333	: orbe_theme( 5, 10, 7)= 1.7833
orbe_theme( 5, 10, 8)= 5.45	 : orbe_theme( 5, 10, 9)= 1.7833	 : orbe_theme( 5, 10, 10)= 1.6333	 : orbe_theme( 5, 10, 11)= 4.3833	 : orbe_theme( 5, 10, 12)= 3.7	 : orbe_theme( 5, 10, 13)= 1.9166	 : orbe_theme( 5, 10, 14)= 1.3	 : orbe_theme( 5, 10, 15)= 0.3
orbe_theme( 5, 11, 0)= 5.45	: orbe_theme( 5, 11, 1)= 0.3	: orbe_theme( 5, 11, 2)= 1.3	: orbe_theme( 5, 11, 3)= 1.9166	: orbe_theme( 5, 11, 4)= 3.7	: orbe_theme( 5, 11, 5)= 4.3833	: orbe_theme( 5, 11, 6)= 1.6333	: orbe_theme( 5, 11, 7)= 1.7833
orbe_theme( 5, 11, 8)= 5.45	 : orbe_theme( 5, 11, 9)= 1.7833	 : orbe_theme( 5, 11, 10)= 1.6333	 : orbe_theme( 5, 11, 11)= 4.3833	 : orbe_theme( 5, 11, 12)= 3.7	 : orbe_theme( 5, 11, 13)= 1.9166	 : orbe_theme( 5, 11, 14)= 1.3	 : orbe_theme( 5, 11, 15)= 0.3


'Saturne, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 6, 0, 0)= 15.7833	: orbe_theme( 6, 0, 1)= 0.7333	: orbe_theme( 6, 0, 2)= 3.1666	: orbe_theme( 6, 0, 3)= 4.75	: orbe_theme( 6, 0, 4)= 9.05	: orbe_theme( 6, 0, 5)= 10.7166	: orbe_theme( 6, 0, 6)= 4.0166	: orbe_theme( 6, 0, 7)= 4.3666
orbe_theme( 6, 0, 8)= 15.7833	 : orbe_theme( 6, 0, 9)= 4.3666	 : orbe_theme( 6, 0, 10)= 4.0166	 : orbe_theme( 6, 0, 11)= 10.7166	 : orbe_theme( 6, 0, 12)= 9.05	 : orbe_theme( 6, 0, 13)= 4.75	 : orbe_theme( 6, 0, 14)= 3.1666	 : orbe_theme( 6, 0, 15)= 0.7333
orbe_theme( 6, 1, 0)= 15.75	: orbe_theme( 6, 1, 1)= 0.7333	: orbe_theme( 6, 1, 2)= 3.1666	: orbe_theme( 6, 1, 3)= 4.7333	: orbe_theme( 6, 1, 4)= 9.0333	: orbe_theme( 6, 1, 5)= 10.6833	: orbe_theme( 6, 1, 6)= 4.0166	: orbe_theme( 6, 1, 7)= 4.35
orbe_theme( 6, 1, 8)= 15.75	 : orbe_theme( 6, 1, 9)= 4.35	 : orbe_theme( 6, 1, 10)= 4.0166	 : orbe_theme( 6, 1, 11)= 10.6833	 : orbe_theme( 6, 1, 12)= 9.0333	 : orbe_theme( 6, 1, 13)= 4.7333	 : orbe_theme( 6, 1, 14)= 3.1666	 : orbe_theme( 6, 1, 15)= 0.7333
orbe_theme( 6, 2, 0)= 8.1166	: orbe_theme( 6, 2, 1)= 0.3666	: orbe_theme( 6, 2, 2)= 1.6333	: orbe_theme( 6, 2, 3)= 2.4166	: orbe_theme( 6, 2, 4)= 4.65	: orbe_theme( 6, 2, 5)= 5.5	: orbe_theme( 6, 2, 6)= 2.0666	: orbe_theme( 6, 2, 7)= 2.2333
orbe_theme( 6, 2, 8)= 8.1166	 : orbe_theme( 6, 2, 9)= 2.2333	 : orbe_theme( 6, 2, 10)= 2.0666	 : orbe_theme( 6, 2, 11)= 5.5	 : orbe_theme( 6, 2, 12)= 4.65	 : orbe_theme( 6, 2, 13)= 2.4166	 : orbe_theme( 6, 2, 14)= 1.6333	 : orbe_theme( 6, 2, 15)= 0.3666
orbe_theme( 6, 3, 0)= 11.8166	: orbe_theme( 6, 3, 1)= 0.55	: orbe_theme( 6, 3, 2)= 2.3833	: orbe_theme( 6, 3, 3)= 3.55	: orbe_theme( 6, 3, 4)= 5.7833	: orbe_theme( 6, 3, 5)= 8.0166	: orbe_theme( 6, 3, 6)= 3	: orbe_theme( 6, 3, 7)= 3.2666
orbe_theme( 6, 3, 8)= 11.8166	 : orbe_theme( 6, 3, 9)= 3.2666	 : orbe_theme( 6, 3, 10)= 3	 : orbe_theme( 6, 3, 11)= 8.0166	 : orbe_theme( 6, 3, 12)= 5.7833	 : orbe_theme( 6, 3, 13)= 3.55	 : orbe_theme( 6, 3, 14)= 2.3833	 : orbe_theme( 6, 3, 15)= 0.55
orbe_theme( 6, 4, 0)= 10.8833	: orbe_theme( 6, 4, 1)= 0.5	: orbe_theme( 6, 4, 2)= 2.1833	: orbe_theme( 6, 4, 3)= 3.25	: orbe_theme( 6, 4, 4)= 5.2333	: orbe_theme( 6, 4, 5)= 7.3833	: orbe_theme( 6, 4, 6)= 2.75	: orbe_theme( 6, 4, 7)= 3
orbe_theme( 6, 4, 8)= 10.8833	 : orbe_theme( 6, 4, 9)= 3	 : orbe_theme( 6, 4, 10)= 2.75	 : orbe_theme( 6, 4, 11)= 7.3833	 : orbe_theme( 6, 4, 12)= 5.2333	 : orbe_theme( 6, 4, 13)= 3.25	 : orbe_theme( 6, 4, 14)= 2.1833	 : orbe_theme( 6, 4, 15)= 0.5
orbe_theme( 6, 5, 0)= 11.9166	: orbe_theme( 6, 5, 1)= 0.55	: orbe_theme( 6, 5, 2)= 2.4	: orbe_theme( 6, 5, 3)= 3.5833	: orbe_theme( 6, 5, 4)= 5.8333	: orbe_theme( 6, 5, 5)= 8.0833	: orbe_theme( 6, 5, 6)= 3.0333	: orbe_theme( 6, 5, 7)= 3.3
orbe_theme( 6, 5, 8)= 11.9166	 : orbe_theme( 6, 5, 9)= 3.3	 : orbe_theme( 6, 5, 10)= 3.0333	 : orbe_theme( 6, 5, 11)= 8.0833	 : orbe_theme( 6, 5, 12)= 5.8333	 : orbe_theme( 6, 5, 13)= 3.5833	 : orbe_theme( 6, 5, 14)= 2.4	 : orbe_theme( 6, 5, 15)= 0.55
orbe_theme( 6, 6, 0)= 10.9166	: orbe_theme( 6, 6, 1)= 0.5	: orbe_theme( 6, 6, 2)= 2.2	: orbe_theme( 6, 6, 3)= 3.2833	: orbe_theme( 6, 6, 4)= 5.25	: orbe_theme( 6, 6, 5)= 7.4166	: orbe_theme( 6, 6, 6)= 2.7833	: orbe_theme( 6, 6, 7)= 3.0166
orbe_theme( 6, 6, 8)= 10.9166	 : orbe_theme( 6, 6, 9)= 3.0166	 : orbe_theme( 6, 6, 10)= 2.7833	 : orbe_theme( 6, 6, 11)= 7.4166	 : orbe_theme( 6, 6, 12)= 5.25	 : orbe_theme( 6, 6, 13)= 3.2833	 : orbe_theme( 6, 6, 14)= 2.2	 : orbe_theme( 6, 6, 15)= 0.5
orbe_theme( 6, 7, 0)= 9.4166	: orbe_theme( 6, 7, 1)= 0.4166	: orbe_theme( 6, 7, 2)= 1.9	: orbe_theme( 6, 7, 3)= 2.8333	: orbe_theme( 6, 7, 4)= 5.4166	: orbe_theme( 6, 7, 5)= 5.4	: orbe_theme( 6, 7, 6)= 2.4	: orbe_theme( 6, 7, 7)= 2.5833
orbe_theme( 6, 7, 8)= 9.4166	 : orbe_theme( 6, 7, 9)= 2.5833	 : orbe_theme( 6, 7, 10)= 2.4	 : orbe_theme( 6, 7, 11)= 5.4	 : orbe_theme( 6, 7, 12)= 5.4166	 : orbe_theme( 6, 7, 13)= 2.8333	 : orbe_theme( 6, 7, 14)= 1.9	 : orbe_theme( 6, 7, 15)= 0.4166
orbe_theme( 6, 8, 0)= 8.9166	: orbe_theme( 6, 8, 1)= 0.4	: orbe_theme( 6, 8, 2)= 1.7833	: orbe_theme( 6, 8, 3)= 2.6833	: orbe_theme( 6, 8, 4)= 5.1166	: orbe_theme( 6, 8, 5)= 5.05	: orbe_theme( 6, 8, 6)= 2.25	: orbe_theme( 6, 8, 7)= 2.4666
orbe_theme( 6, 8, 8)= 8.9166	 : orbe_theme( 6, 8, 9)= 2.4666	 : orbe_theme( 6, 8, 10)= 2.25	 : orbe_theme( 6, 8, 11)= 5.05	 : orbe_theme( 6, 8, 12)= 5.1166	 : orbe_theme( 6, 8, 13)= 2.6833	 : orbe_theme( 6, 8, 14)= 1.7833	 : orbe_theme( 6, 8, 15)= 0.4
orbe_theme( 6, 9, 0)= 8.8	: orbe_theme( 6, 9, 1)= 0.4	: orbe_theme( 6, 9, 2)= 1.75	: orbe_theme( 6, 9, 3)= 2.65	: orbe_theme( 6, 9, 4)= 5.05	: orbe_theme( 6, 9, 5)= 5.9666	: orbe_theme( 6, 9, 6)= 2.2333	: orbe_theme( 6, 9, 7)= 2.4166
orbe_theme( 6, 9, 8)= 8.8	 : orbe_theme( 6, 9, 9)= 2.4166	 : orbe_theme( 6, 9, 10)= 2.2333	 : orbe_theme( 6, 9, 11)= 5.9666	 : orbe_theme( 6, 9, 12)= 5.05	 : orbe_theme( 6, 9, 13)= 2.65	 : orbe_theme( 6, 9, 14)= 1.75	 : orbe_theme( 6, 9, 15)= 0.4
orbe_theme( 6, 10, 0)= 5.45	: orbe_theme( 6, 10, 1)= 0.25	: orbe_theme( 6, 10, 2)= 1.0833	: orbe_theme( 6, 10, 3)= 1.6333	: orbe_theme( 6, 10, 4)= 3.1333	: orbe_theme( 6, 10, 5)= 3.7	: orbe_theme( 6, 10, 6)= 1.3833	: orbe_theme( 6, 10, 7)= 1.5
orbe_theme( 6, 10, 8)= 5.45	 : orbe_theme( 6, 10, 9)= 1.5	 : orbe_theme( 6, 10, 10)= 1.3833	 : orbe_theme( 6, 10, 11)= 3.7	 : orbe_theme( 6, 10, 12)= 3.1333	 : orbe_theme( 6, 10, 13)= 1.6333	 : orbe_theme( 6, 10, 14)= 1.0833	 : orbe_theme( 6, 10, 15)= 0.25
orbe_theme( 6, 11, 0)= 5.45	: orbe_theme( 6, 11, 1)= 0.25	: orbe_theme( 6, 11, 2)= 1.0833	: orbe_theme( 6, 11, 3)= 1.6333	: orbe_theme( 6, 11, 4)= 3.1333	: orbe_theme( 6, 11, 5)= 3.7	: orbe_theme( 6, 11, 6)= 1.3833	: orbe_theme( 6, 11, 7)= 1.5
orbe_theme( 6, 11, 8)= 5.45	 : orbe_theme( 6, 11, 9)= 1.5	 : orbe_theme( 6, 11, 10)= 1.3833	 : orbe_theme( 6, 11, 11)= 3.7	 : orbe_theme( 6, 11, 12)= 3.1333	 : orbe_theme( 6, 11, 13)= 1.6333	 : orbe_theme( 6, 11, 14)= 1.0833	 : orbe_theme( 6, 11, 15)= 0.25


'Uranus, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 7, 0, 0)= 14.3	: orbe_theme( 7, 0, 1)= 0.6666	: orbe_theme( 7, 0, 2)= 2.8833	: orbe_theme( 7, 0, 3)= 4.3	: orbe_theme( 7, 0, 4)= 8.2	: orbe_theme( 7, 0, 5)= 9.7	: orbe_theme( 7, 0, 6)= 3.6333	: orbe_theme( 7, 0, 7)= 3.95
orbe_theme( 7, 0, 8)= 14.3	 : orbe_theme( 7, 0, 9)= 3.95	 : orbe_theme( 7, 0, 10)= 3.6333	 : orbe_theme( 7, 0, 11)= 9.7	 : orbe_theme( 7, 0, 12)= 8.2	 : orbe_theme( 7, 0, 13)= 4.3	 : orbe_theme( 7, 0, 14)= 2.8833	 : orbe_theme( 7, 0, 15)= 0.6666
orbe_theme( 7, 1, 0)= 14.2666	: orbe_theme( 7, 1, 1)= 0.65	: orbe_theme( 7, 1, 2)= 2.8666	: orbe_theme( 7, 1, 3)= 4.2833	: orbe_theme( 7, 1, 4)= 8.1833	: orbe_theme( 7, 1, 5)= 9.6833	: orbe_theme( 7, 1, 6)= 3.6333	: orbe_theme( 7, 1, 7)= 3.95
orbe_theme( 7, 1, 8)= 14.2666	 : orbe_theme( 7, 1, 9)= 3.95	 : orbe_theme( 7, 1, 10)= 3.6333	 : orbe_theme( 7, 1, 11)= 9.6833	 : orbe_theme( 7, 1, 12)= 8.1833	 : orbe_theme( 7, 1, 13)= 4.2833	 : orbe_theme( 7, 1, 14)= 2.8666	 : orbe_theme( 7, 1, 15)= 0.65
orbe_theme( 7, 2, 0)= 5.6333	: orbe_theme( 7, 2, 1)= 0.3	: orbe_theme( 7, 2, 2)= 1.3333	: orbe_theme( 7, 2, 3)= 2	: orbe_theme( 7, 2, 4)= 3.8	: orbe_theme( 7, 2, 5)= 4.5	: orbe_theme( 7, 2, 6)= 1.6833	: orbe_theme( 7, 2, 7)= 1.8333
orbe_theme( 7, 2, 8)= 5.6333	 : orbe_theme( 7, 2, 9)= 1.8333	 : orbe_theme( 7, 2, 10)= 1.6833	 : orbe_theme( 7, 2, 11)= 4.5	 : orbe_theme( 7, 2, 12)= 3.8	 : orbe_theme( 7, 2, 13)= 2	 : orbe_theme( 7, 2, 14)= 1.3333	 : orbe_theme( 7, 2, 15)= 0.3
orbe_theme( 7, 3, 0)= 10.35	: orbe_theme( 7, 3, 1)= 0.4666	: orbe_theme( 7, 3, 2)= 2.0833	: orbe_theme( 7, 3, 3)= 3.1166	: orbe_theme( 7, 3, 4)= 5.9166	: orbe_theme( 7, 3, 5)= 7.0166	: orbe_theme( 7, 3, 6)= 2.6333	: orbe_theme( 7, 3, 7)= 2.8666
orbe_theme( 7, 3, 8)= 10.35	 : orbe_theme( 7, 3, 9)= 2.8666	 : orbe_theme( 7, 3, 10)= 2.6333	 : orbe_theme( 7, 3, 11)= 7.0166	 : orbe_theme( 7, 3, 12)= 5.9166	 : orbe_theme( 7, 3, 13)= 3.1166	 : orbe_theme( 7, 3, 14)= 2.0833	 : orbe_theme( 7, 3, 15)= 0.4666
orbe_theme( 7, 4, 0)= 9.4	: orbe_theme( 7, 4, 1)= 0.4166	: orbe_theme( 7, 4, 2)= 1.8833	: orbe_theme( 7, 4, 3)= 2.8166	: orbe_theme( 7, 4, 4)= 5.4	: orbe_theme( 7, 4, 5)= 5.3833	: orbe_theme( 7, 4, 6)= 2.3833	: orbe_theme( 7, 4, 7)= 2.5833
orbe_theme( 7, 4, 8)= 9.4	 : orbe_theme( 7, 4, 9)= 2.5833	 : orbe_theme( 7, 4, 10)= 2.3833	 : orbe_theme( 7, 4, 11)= 5.3833	 : orbe_theme( 7, 4, 12)= 5.4	 : orbe_theme( 7, 4, 13)= 2.8166	 : orbe_theme( 7, 4, 14)= 1.8833	 : orbe_theme( 7, 4, 15)= 0.4166
orbe_theme( 7, 5, 0)= 10.4166	: orbe_theme( 7, 5, 1)= 0.4833	: orbe_theme( 7, 5, 2)= 2.0833	: orbe_theme( 7, 5, 3)= 3.1333	: orbe_theme( 7, 5, 4)= 5.9833	: orbe_theme( 7, 5, 5)= 7.0833	: orbe_theme( 7, 5, 6)= 2.65	: orbe_theme( 7, 5, 7)= 2.8833
orbe_theme( 7, 5, 8)= 10.4166	 : orbe_theme( 7, 5, 9)= 2.8833	 : orbe_theme( 7, 5, 10)= 2.65	 : orbe_theme( 7, 5, 11)= 7.0833	 : orbe_theme( 7, 5, 12)= 5.9833	 : orbe_theme( 7, 5, 13)= 3.1333	 : orbe_theme( 7, 5, 14)= 2.0833	 : orbe_theme( 7, 5, 15)= 0.4833
orbe_theme( 7, 6, 0)= 9.4166	: orbe_theme( 7, 6, 1)= 0.4166	: orbe_theme( 7, 6, 2)= 1.9	: orbe_theme( 7, 6, 3)= 2.8333	: orbe_theme( 7, 6, 4)= 5.4166	: orbe_theme( 7, 6, 5)= 5.4	: orbe_theme( 7, 6, 6)= 2.4	: orbe_theme( 7, 6, 7)= 2.5833
orbe_theme( 7, 6, 8)= 9.4166	 : orbe_theme( 7, 6, 9)= 2.5833	 : orbe_theme( 7, 6, 10)= 2.4	 : orbe_theme( 7, 6, 11)= 5.4	 : orbe_theme( 7, 6, 12)= 5.4166	 : orbe_theme( 7, 6, 13)= 2.8333	 : orbe_theme( 7, 6, 14)= 1.9	 : orbe_theme( 7, 6, 15)= 0.4166
orbe_theme( 7, 7, 0)= 7.9666	: orbe_theme( 7, 7, 1)= 0.3666	: orbe_theme( 7, 7, 2)= 1.5833	: orbe_theme( 7, 7, 3)= 2.4	: orbe_theme( 7, 7, 4)= 4.5666	: orbe_theme( 7, 7, 5)= 5.4	: orbe_theme( 7, 7, 6)= 2.0166	: orbe_theme( 7, 7, 7)= 2.2
orbe_theme( 7, 7, 8)= 7.9666	 : orbe_theme( 7, 7, 9)= 2.2	 : orbe_theme( 7, 7, 10)= 2.0166	 : orbe_theme( 7, 7, 11)= 5.4	 : orbe_theme( 7, 7, 12)= 4.5666	 : orbe_theme( 7, 7, 13)= 2.4	 : orbe_theme( 7, 7, 14)= 1.5833	 : orbe_theme( 7, 7, 15)= 0.3666
orbe_theme( 7, 8, 0)= 7.4166	: orbe_theme( 7, 8, 1)= 0.3333	: orbe_theme( 7, 8, 2)= 1.5	: orbe_theme( 7, 8, 3)= 2.2333	: orbe_theme( 7, 8, 4)= 4.25	: orbe_theme( 7, 8, 5)= 5.05	: orbe_theme( 7, 8, 6)= 1.8833	: orbe_theme( 7, 8, 7)= 2.05
orbe_theme( 7, 8, 8)= 7.4166	 : orbe_theme( 7, 8, 9)= 2.05	 : orbe_theme( 7, 8, 10)= 1.8833	 : orbe_theme( 7, 8, 11)= 5.05	 : orbe_theme( 7, 8, 12)= 4.25	 : orbe_theme( 7, 8, 13)= 2.2333	 : orbe_theme( 7, 8, 14)= 1.5	 : orbe_theme( 7, 8, 15)= 0.3333
orbe_theme( 7, 9, 0)= 7.3333	: orbe_theme( 7, 9, 1)= 0.3333	: orbe_theme( 7, 9, 2)= 1.4666	: orbe_theme( 7, 9, 3)= 2.2	: orbe_theme( 7, 9, 4)= 4.2	: orbe_theme( 7, 9, 5)= 4.9666	: orbe_theme( 7, 9, 6)= 1.8666	: orbe_theme( 7, 9, 7)= 2.0166
orbe_theme( 7, 9, 8)= 7.3333	 : orbe_theme( 7, 9, 9)= 2.0166	 : orbe_theme( 7, 9, 10)= 1.8666	 : orbe_theme( 7, 9, 11)= 4.9666	 : orbe_theme( 7, 9, 12)= 4.2	 : orbe_theme( 7, 9, 13)= 2.2	 : orbe_theme( 7, 9, 14)= 1.4666	 : orbe_theme( 7, 9, 15)= 0.3333
orbe_theme( 7, 10, 0)= 3.9833	: orbe_theme( 7, 10, 1)= 0.1833	: orbe_theme( 7, 10, 2)= 0.8	: orbe_theme( 7, 10, 3)= 1.2	: orbe_theme( 7, 10, 4)= 2.2833	: orbe_theme( 7, 10, 5)= 2.7	: orbe_theme( 7, 10, 6)= 1	: orbe_theme( 7, 10, 7)= 1.0833
orbe_theme( 7, 10, 8)= 3.9833	 : orbe_theme( 7, 10, 9)= 1.0833	 : orbe_theme( 7, 10, 10)= 1	 : orbe_theme( 7, 10, 11)= 2.7	 : orbe_theme( 7, 10, 12)= 2.2833	 : orbe_theme( 7, 10, 13)= 1.2	 : orbe_theme( 7, 10, 14)= 0.8	 : orbe_theme( 7, 10, 15)= 0.1833
orbe_theme( 7, 11, 0)= 3.9833	: orbe_theme( 7, 11, 1)= 0.1833	: orbe_theme( 7, 11, 2)= 0.8	: orbe_theme( 7, 11, 3)= 1.2	: orbe_theme( 7, 11, 4)= 2.2833	: orbe_theme( 7, 11, 5)= 2.7	: orbe_theme( 7, 11, 6)= 1	: orbe_theme( 7, 11, 7)= 1.0833
orbe_theme( 7, 11, 8)= 3.9833	 : orbe_theme( 7, 11, 9)= 1.0833	 : orbe_theme( 7, 11, 10)= 1	 : orbe_theme( 7, 11, 11)= 2.7	 : orbe_theme( 7, 11, 12)= 2.2833	 : orbe_theme( 7, 11, 13)= 1.2	 : orbe_theme( 7, 11, 14)= 0.8	 : orbe_theme( 7, 11, 15)= 0.1833


'Neptune, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 8, 0, 0)= 13.7666	: orbe_theme( 8, 0, 1)= 0.6333	: orbe_theme( 8, 0, 2)= 2.7666	: orbe_theme( 8, 0, 3)= 4.1333	: orbe_theme( 8, 0, 4)= 7.9	: orbe_theme( 8, 0, 5)= 9.35	: orbe_theme( 8, 0, 6)= 3.5	: orbe_theme( 8, 0, 7)= 3.8
orbe_theme( 8, 0, 8)= 13.7666	 : orbe_theme( 8, 0, 9)= 3.8	 : orbe_theme( 8, 0, 10)= 3.5	 : orbe_theme( 8, 0, 11)= 9.35	 : orbe_theme( 8, 0, 12)= 7.9	 : orbe_theme( 8, 0, 13)= 4.1333	 : orbe_theme( 8, 0, 14)= 2.7666	 : orbe_theme( 8, 0, 15)= 0.6333
orbe_theme( 8, 1, 0)= 13.75	: orbe_theme( 8, 1, 1)= 0.6333	: orbe_theme( 8, 1, 2)= 2.7666	: orbe_theme( 8, 1, 3)= 4.1333	: orbe_theme( 8, 1, 4)= 7.8833	: orbe_theme( 8, 1, 5)= 9.3333	: orbe_theme( 8, 1, 6)= 3.5	: orbe_theme( 8, 1, 7)= 3.8
orbe_theme( 8, 1, 8)= 13.75	 : orbe_theme( 8, 1, 9)= 3.8	 : orbe_theme( 8, 1, 10)= 3.5	 : orbe_theme( 8, 1, 11)= 9.3333	 : orbe_theme( 8, 1, 12)= 7.8833	 : orbe_theme( 8, 1, 13)= 4.1333	 : orbe_theme( 8, 1, 14)= 2.7666	 : orbe_theme( 8, 1, 15)= 0.6333
orbe_theme( 8, 2, 0)= 6.1	: orbe_theme( 8, 2, 1)= 0.2833	: orbe_theme( 8, 2, 2)= 1.2166	: orbe_theme( 8, 2, 3)= 1.8333	: orbe_theme( 8, 2, 4)= 3.5	: orbe_theme( 8, 2, 5)= 4.15	: orbe_theme( 8, 2, 6)= 1.55	: orbe_theme( 8, 2, 7)= 1.6833
orbe_theme( 8, 2, 8)= 6.1	 : orbe_theme( 8, 2, 9)= 1.6833	 : orbe_theme( 8, 2, 10)= 1.55	 : orbe_theme( 8, 2, 11)= 4.15	 : orbe_theme( 8, 2, 12)= 3.5	 : orbe_theme( 8, 2, 13)= 1.8333	 : orbe_theme( 8, 2, 14)= 1.2166	 : orbe_theme( 8, 2, 15)= 0.2833
orbe_theme( 8, 3, 0)= 9.8166	: orbe_theme( 8, 3, 1)= 0.45	: orbe_theme( 8, 3, 2)= 1.9666	: orbe_theme( 8, 3, 3)= 2.95	: orbe_theme( 8, 3, 4)= 5.6333	: orbe_theme( 8, 3, 5)= 6.6666	: orbe_theme( 8, 3, 6)= 2.5	: orbe_theme( 8, 3, 7)= 2.7166
orbe_theme( 8, 3, 8)= 9.8166	 : orbe_theme( 8, 3, 9)= 2.7166	 : orbe_theme( 8, 3, 10)= 2.5	 : orbe_theme( 8, 3, 11)= 6.6666	 : orbe_theme( 8, 3, 12)= 5.6333	 : orbe_theme( 8, 3, 13)= 2.95	 : orbe_theme( 8, 3, 14)= 1.9666	 : orbe_theme( 8, 3, 15)= 0.45
orbe_theme( 8, 4, 0)= 8.8666	: orbe_theme( 8, 4, 1)= 0.4	: orbe_theme( 8, 4, 2)= 1.7833	: orbe_theme( 8, 4, 3)= 2.6666	: orbe_theme( 8, 4, 4)= 5.0833	: orbe_theme( 8, 4, 5)= 6.0166	: orbe_theme( 8, 4, 6)= 2.25	: orbe_theme( 8, 4, 7)= 2.45
orbe_theme( 8, 4, 8)= 8.8666	 : orbe_theme( 8, 4, 9)= 2.45	 : orbe_theme( 8, 4, 10)= 2.25	 : orbe_theme( 8, 4, 11)= 6.0166	 : orbe_theme( 8, 4, 12)= 5.0833	 : orbe_theme( 8, 4, 13)= 2.6666	 : orbe_theme( 8, 4, 14)= 1.7833	 : orbe_theme( 8, 4, 15)= 0.4
orbe_theme( 8, 5, 0)= 9.9	: orbe_theme( 8, 5, 1)= 0.45	: orbe_theme( 8, 5, 2)= 1.9833	: orbe_theme( 8, 5, 3)= 2.9833	: orbe_theme( 8, 5, 4)= 5.6833	: orbe_theme( 8, 5, 5)= 6.7166	: orbe_theme( 8, 5, 6)= 2.5166	: orbe_theme( 8, 5, 7)= 2.7333
orbe_theme( 8, 5, 8)= 9.9	 : orbe_theme( 8, 5, 9)= 2.7333	 : orbe_theme( 8, 5, 10)= 2.5166	 : orbe_theme( 8, 5, 11)= 6.7166	 : orbe_theme( 8, 5, 12)= 5.6833	 : orbe_theme( 8, 5, 13)= 2.9833	 : orbe_theme( 8, 5, 14)= 1.9833	 : orbe_theme( 8, 5, 15)= 0.45
orbe_theme( 8, 6, 0)= 8.9166	: orbe_theme( 8, 6, 1)= 0.4	: orbe_theme( 8, 6, 2)= 1.7833	: orbe_theme( 8, 6, 3)= 2.6833	: orbe_theme( 8, 6, 4)= 5.1166	: orbe_theme( 8, 6, 5)= 6.05	: orbe_theme( 8, 6, 6)= 2.2666	: orbe_theme( 8, 6, 7)= 2.4666
orbe_theme( 8, 6, 8)= 8.9166	 : orbe_theme( 8, 6, 9)= 2.4666	 : orbe_theme( 8, 6, 10)= 2.2666	 : orbe_theme( 8, 6, 11)= 6.05	 : orbe_theme( 8, 6, 12)= 5.1166	 : orbe_theme( 8, 6, 13)= 2.6833	 : orbe_theme( 8, 6, 14)= 1.7833	 : orbe_theme( 8, 6, 15)= 0.4
orbe_theme( 8, 7, 0)= 7.4333	: orbe_theme( 8, 7, 1)= 0.3333	: orbe_theme( 8, 7, 2)= 1.5	: orbe_theme( 8, 7, 3)= 2.2333	: orbe_theme( 8, 7, 4)= 4.2666	: orbe_theme( 8, 7, 5)= 5.05	: orbe_theme( 8, 7, 6)= 1.8833	: orbe_theme( 8, 7, 7)= 2.05
orbe_theme( 8, 7, 8)= 7.4333	 : orbe_theme( 8, 7, 9)= 2.05	 : orbe_theme( 8, 7, 10)= 1.8833	 : orbe_theme( 8, 7, 11)= 5.05	 : orbe_theme( 8, 7, 12)= 4.2666	 : orbe_theme( 8, 7, 13)= 2.2333	 : orbe_theme( 8, 7, 14)= 1.5	 : orbe_theme( 8, 7, 15)= 0.3333
orbe_theme( 8, 8, 0)= 6.9	: orbe_theme( 8, 8, 1)= 0.3166	: orbe_theme( 8, 8, 2)= 1.3833	: orbe_theme( 8, 8, 3)= 2.0666	: orbe_theme( 8, 8, 4)= 3.9666	: orbe_theme( 8, 8, 5)= 4.6833	: orbe_theme( 8, 8, 6)= 1.75	: orbe_theme( 8, 8, 7)= 1.9
orbe_theme( 8, 8, 8)= 6.9	 : orbe_theme( 8, 8, 9)= 1.9	 : orbe_theme( 8, 8, 10)= 1.75	 : orbe_theme( 8, 8, 11)= 4.6833	 : orbe_theme( 8, 8, 12)= 3.9666	 : orbe_theme( 8, 8, 13)= 2.0666	 : orbe_theme( 8, 8, 14)= 1.3833	 : orbe_theme( 8, 8, 15)= 0.3166
orbe_theme( 8, 9, 0)= 6.8	: orbe_theme( 8, 9, 1)= 0.3166	: orbe_theme( 8, 9, 2)= 1.3666	: orbe_theme( 8, 9, 3)= 2.0333	: orbe_theme( 8, 9, 4)= 3.9	: orbe_theme( 8, 9, 5)= 4.6166	: orbe_theme( 8, 9, 6)= 1.7333	: orbe_theme( 8, 9, 7)= 1.8833
orbe_theme( 8, 9, 8)= 6.8	 : orbe_theme( 8, 9, 9)= 1.8833	 : orbe_theme( 8, 9, 10)= 1.7333	 : orbe_theme( 8, 9, 11)= 4.6166	 : orbe_theme( 8, 9, 12)= 3.9	 : orbe_theme( 8, 9, 13)= 2.0333	 : orbe_theme( 8, 9, 14)= 1.3666	 : orbe_theme( 8, 9, 15)= 0.3166
orbe_theme( 8, 10, 0)= 3.45	: orbe_theme( 8, 10, 1)= 0.15	: orbe_theme( 8, 10, 2)= 0.6833	: orbe_theme( 8, 10, 3)= 1.0333	: orbe_theme( 8, 10, 4)= 1.9833	: orbe_theme( 8, 10, 5)= 2.3333	: orbe_theme( 8, 10, 6)= 0.8666	: orbe_theme( 8, 10, 7)= 0.95
orbe_theme( 8, 10, 8)= 3.45	 : orbe_theme( 8, 10, 9)= 0.95	 : orbe_theme( 8, 10, 10)= 0.8666	 : orbe_theme( 8, 10, 11)= 2.3333	 : orbe_theme( 8, 10, 12)= 1.9833	 : orbe_theme( 8, 10, 13)= 1.0333	 : orbe_theme( 8, 10, 14)= 0.6833	 : orbe_theme( 8, 10, 15)= 0.15
orbe_theme( 8, 11, 0)= 3.45	: orbe_theme( 8, 11, 1)= 0.15	: orbe_theme( 8, 11, 2)= 0.6833	: orbe_theme( 8, 11, 3)= 1.0333	: orbe_theme( 8, 11, 4)= 1.9833	: orbe_theme( 8, 11, 5)= 2.3333	: orbe_theme( 8, 11, 6)= 0.8666	: orbe_theme( 8, 11, 7)= 0.95
orbe_theme( 8, 11, 8)= 3.45	 : orbe_theme( 8, 11, 9)= 0.95	 : orbe_theme( 8, 11, 10)= 0.8666	 : orbe_theme( 8, 11, 11)= 2.3333	 : orbe_theme( 8, 11, 12)= 1.9833	 : orbe_theme( 8, 11, 13)= 1.0333	 : orbe_theme( 8, 11, 14)= 0.6833	 : orbe_theme( 8, 11, 15)= 0.15


'Pluton, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 9, 0, 0)= 13.6666	: orbe_theme( 9, 0, 1)= 0.6333	: orbe_theme( 9, 0, 2)= 2.75	: orbe_theme( 9, 0, 3)= 4.1	: orbe_theme( 9, 0, 4)= 7.8333	: orbe_theme( 9, 0, 5)= 9.2666	: orbe_theme( 9, 0, 6)= 3.4833	: orbe_theme( 9, 0, 7)= 3.7833
orbe_theme( 9, 0, 8)= 13.6666	 : orbe_theme( 9, 0, 9)= 3.7833	 : orbe_theme( 9, 0, 10)= 3.4833	 : orbe_theme( 9, 0, 11)= 9.2666	 : orbe_theme( 9, 0, 12)= 7.8333	 : orbe_theme( 9, 0, 13)= 4.1	 : orbe_theme( 9, 0, 14)= 2.75	 : orbe_theme( 9, 0, 15)= 0.6333
orbe_theme( 9, 1, 0)= 13.6333	: orbe_theme( 9, 1, 1)= 0.6333	: orbe_theme( 9, 1, 2)= 2.75	: orbe_theme( 9, 1, 3)= 4.1	: orbe_theme( 9, 1, 4)= 7.8166	: orbe_theme( 9, 1, 5)= 9.25	: orbe_theme( 9, 1, 6)= 3.4666	: orbe_theme( 9, 1, 7)= 3.7666
orbe_theme( 9, 1, 8)= 13.6333	 : orbe_theme( 9, 1, 9)= 3.7666	 : orbe_theme( 9, 1, 10)= 3.4666	 : orbe_theme( 9, 1, 11)= 9.25	 : orbe_theme( 9, 1, 12)= 7.8166	 : orbe_theme( 9, 1, 13)= 4.1	 : orbe_theme( 9, 1, 14)= 2.75	 : orbe_theme( 9, 1, 15)= 0.6333
orbe_theme( 9, 2, 0)= 6	: orbe_theme( 9, 2, 1)= 0.2666	: orbe_theme( 9, 2, 2)= 1.2	: orbe_theme( 9, 2, 3)= 1.8	: orbe_theme( 9, 2, 4)= 3.4333	: orbe_theme( 9, 2, 5)= 4.0666	: orbe_theme( 9, 2, 6)= 1.5166	: orbe_theme( 9, 2, 7)= 1.65
orbe_theme( 9, 2, 8)= 6	 : orbe_theme( 9, 2, 9)= 1.65	 : orbe_theme( 9, 2, 10)= 1.5166	 : orbe_theme( 9, 2, 11)= 4.0666	 : orbe_theme( 9, 2, 12)= 3.4333	 : orbe_theme( 9, 2, 13)= 1.8	 : orbe_theme( 9, 2, 14)= 1.2	 : orbe_theme( 9, 2, 15)= 0.2666
orbe_theme( 9, 3, 0)= 9.7166	: orbe_theme( 9, 3, 1)= 0.45	: orbe_theme( 9, 3, 2)= 1.95	: orbe_theme( 9, 3, 3)= 2.9166	: orbe_theme( 9, 3, 4)= 5.5666	: orbe_theme( 9, 3, 5)= 6.5833	: orbe_theme( 9, 3, 6)= 2.4666	: orbe_theme( 9, 3, 7)= 2.6833
orbe_theme( 9, 3, 8)= 9.7166	 : orbe_theme( 9, 3, 9)= 2.6833	 : orbe_theme( 9, 3, 10)= 2.4666	 : orbe_theme( 9, 3, 11)= 6.5833	 : orbe_theme( 9, 3, 12)= 5.5666	 : orbe_theme( 9, 3, 13)= 2.9166	 : orbe_theme( 9, 3, 14)= 1.95	 : orbe_theme( 9, 3, 15)= 0.45
orbe_theme( 9, 4, 0)= 8.7666	: orbe_theme( 9, 4, 1)= 0.4	: orbe_theme( 9, 4, 2)= 1.7666	: orbe_theme( 9, 4, 3)= 2.6333	: orbe_theme( 9, 4, 4)= 5.0333	: orbe_theme( 9, 4, 5)= 5.95	: orbe_theme( 9, 4, 6)= 2.2333	: orbe_theme( 9, 4, 7)= 2.4166
orbe_theme( 9, 4, 8)= 8.7666	 : orbe_theme( 9, 4, 9)= 2.4166	 : orbe_theme( 9, 4, 10)= 2.2333	 : orbe_theme( 9, 4, 11)= 5.95	 : orbe_theme( 9, 4, 12)= 5.0333	 : orbe_theme( 9, 4, 13)= 2.6333	 : orbe_theme( 9, 4, 14)= 1.7666	 : orbe_theme( 9, 4, 15)= 0.4
orbe_theme( 9, 5, 0)= 9.8	: orbe_theme( 9, 5, 1)= 0.45	: orbe_theme( 9, 5, 2)= 1.9666	: orbe_theme( 9, 5, 3)= 2.95	: orbe_theme( 9, 5, 4)= 5.6166	: orbe_theme( 9, 5, 5)= 6.65	: orbe_theme( 9, 5, 6)= 2.5	: orbe_theme( 9, 5, 7)= 2.7
orbe_theme( 9, 5, 8)= 9.8	 : orbe_theme( 9, 5, 9)= 2.7	 : orbe_theme( 9, 5, 10)= 2.5	 : orbe_theme( 9, 5, 11)= 6.65	 : orbe_theme( 9, 5, 12)= 5.6166	 : orbe_theme( 9, 5, 13)= 2.95	 : orbe_theme( 9, 5, 14)= 1.9666	 : orbe_theme( 9, 5, 15)= 0.45
orbe_theme( 9, 6, 0)= 8.8	: orbe_theme( 9, 6, 1)= 0.4	: orbe_theme( 9, 6, 2)= 1.7666	: orbe_theme( 9, 6, 3)= 2.65	: orbe_theme( 9, 6, 4)= 5.05	: orbe_theme( 9, 6, 5)= 5.9666	: orbe_theme( 9, 6, 6)= 2.2333	: orbe_theme( 9, 6, 7)= 2.4333
orbe_theme( 9, 6, 8)= 8.8	 : orbe_theme( 9, 6, 9)= 2.4333	 : orbe_theme( 9, 6, 10)= 2.2333	 : orbe_theme( 9, 6, 11)= 5.9666	 : orbe_theme( 9, 6, 12)= 5.05	 : orbe_theme( 9, 6, 13)= 2.65	 : orbe_theme( 9, 6, 14)= 1.7666	 : orbe_theme( 9, 6, 15)= 0.4
orbe_theme( 9, 7, 0)= 7.3333	: orbe_theme( 9, 7, 1)= 0.3333	: orbe_theme( 9, 7, 2)= 1.4666	: orbe_theme( 9, 7, 3)= 2.2	: orbe_theme( 9, 7, 4)= 4.2	: orbe_theme( 9, 7, 5)= 4.9666	: orbe_theme( 9, 7, 6)= 1.8666	: orbe_theme( 9, 7, 7)= 2.0166
orbe_theme( 9, 7, 8)= 7.3333	 : orbe_theme( 9, 7, 9)= 2.0166	 : orbe_theme( 9, 7, 10)= 1.8666	 : orbe_theme( 9, 7, 11)= 4.9666	 : orbe_theme( 9, 7, 12)= 4.2	 : orbe_theme( 9, 7, 13)= 2.2	 : orbe_theme( 9, 7, 14)= 1.4666	 : orbe_theme( 9, 7, 15)= 0.3333
orbe_theme( 9, 8, 0)= 6.8	: orbe_theme( 9, 8, 1)= 0.3166	: orbe_theme( 9, 8, 2)= 1.3666	: orbe_theme( 9, 8, 3)= 2.0333	: orbe_theme( 9, 8, 4)= 3.9	: orbe_theme( 9, 8, 5)= 4.6166	: orbe_theme( 9, 8, 6)= 1.7333	: orbe_theme( 9, 8, 7)= 1.8833
orbe_theme( 9, 8, 8)= 6.8	 : orbe_theme( 9, 8, 9)= 1.8833	 : orbe_theme( 9, 8, 10)= 1.7333	 : orbe_theme( 9, 8, 11)= 4.6166	 : orbe_theme( 9, 8, 12)= 3.9	 : orbe_theme( 9, 8, 13)= 2.0333	 : orbe_theme( 9, 8, 14)= 1.3666	 : orbe_theme( 9, 8, 15)= 0.3166
orbe_theme( 9, 9, 0)= 6.6833	: orbe_theme( 9, 9, 1)= 0.3	: orbe_theme( 9, 9, 2)= 1.35	: orbe_theme( 9, 9, 3)= 2	: orbe_theme( 9, 9, 4)= 3.8333	: orbe_theme( 9, 9, 5)= 4.5333	: orbe_theme( 9, 9, 6)= 1.7	: orbe_theme( 9, 9, 7)= 1.85
orbe_theme( 9, 9, 8)= 6.6833	 : orbe_theme( 9, 9, 9)= 1.85	 : orbe_theme( 9, 9, 10)= 1.7	 : orbe_theme( 9, 9, 11)= 4.5333	 : orbe_theme( 9, 9, 12)= 3.8333	 : orbe_theme( 9, 9, 13)= 2	 : orbe_theme( 9, 9, 14)= 1.35	 : orbe_theme( 9, 9, 15)= 0.3
orbe_theme( 9, 10, 0)= 3.3333	: orbe_theme( 9, 10, 1)= 0.15	: orbe_theme( 9, 10, 2)= 0.6666	: orbe_theme( 9, 10, 3)= 1	: orbe_theme( 9, 10, 4)= 1.9166	: orbe_theme( 9, 10, 5)= 2.2666	: orbe_theme( 9, 10, 6)= 0.85	: orbe_theme( 9, 10, 7)= 0.9166
orbe_theme( 9, 10, 8)= 3.3333	 : orbe_theme( 9, 10, 9)= 0.9166	 : orbe_theme( 9, 10, 10)= 0.85	 : orbe_theme( 9, 10, 11)= 2.2666	 : orbe_theme( 9, 10, 12)= 1.9166	 : orbe_theme( 9, 10, 13)= 1	 : orbe_theme( 9, 10, 14)= 0.6666	 : orbe_theme( 9, 10, 15)= 0.15
orbe_theme( 9, 11, 0)= 3.3333	: orbe_theme( 9, 11, 1)= 0.15	: orbe_theme( 9, 11, 2)= 0.6666	: orbe_theme( 9, 11, 3)= 1	: orbe_theme( 9, 11, 4)= 1.9166	: orbe_theme( 9, 11, 5)= 2.2666	: orbe_theme( 9, 11, 6)= 0.85	: orbe_theme( 9, 11, 7)= 0.9166
orbe_theme( 9, 11, 8)= 3.3333	 : orbe_theme( 9, 11, 9)= 0.9166	 : orbe_theme( 9, 11, 10)= 0.85	 : orbe_theme( 9, 11, 11)= 2.2666	 : orbe_theme( 9, 11, 12)= 1.9166	 : orbe_theme( 9, 11, 13)= 1	 : orbe_theme( 9, 11, 14)= 0.6666	 : orbe_theme( 9, 11, 15)= 0.15


'Noeuds lunaires, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 10, 0, 0)= 10.3166	: orbe_theme( 10, 0, 1)= 0.4666	: orbe_theme( 10, 0, 2)= 2.0666	: orbe_theme( 10, 0, 3)= 3.1	: orbe_theme( 10, 0, 4)= 5.9166	: orbe_theme( 10, 0, 5)= 7	: orbe_theme( 10, 0, 6)= 2.6166	: orbe_theme( 10, 0, 7)= 2.85
orbe_theme( 10, 0, 8)= 10.3166	 : orbe_theme( 10, 0, 9)= 2.85	 : orbe_theme( 10, 0, 10)= 2.6166	 : orbe_theme( 10, 0, 11)= 7	 : orbe_theme( 10, 0, 12)= 5.9166	 : orbe_theme( 10, 0, 13)= 3.1	 : orbe_theme( 10, 0, 14)= 2.0666	 : orbe_theme( 10, 0, 15)= 0.4666
orbe_theme( 10, 1, 0)= 10.2833	: orbe_theme( 10, 1, 1)= 0.4666	: orbe_theme( 10, 1, 2)= 2.0666	: orbe_theme( 10, 1, 3)= 3.0833	: orbe_theme( 10, 1, 4)= 5.9	: orbe_theme( 10, 1, 5)= 6.9833	: orbe_theme( 10, 1, 6)= 2.6166	: orbe_theme( 10, 1, 7)= 2.85
orbe_theme( 10, 1, 8)= 10.2833	 : orbe_theme( 10, 1, 9)= 2.85	 : orbe_theme( 10, 1, 10)= 2.6166	 : orbe_theme( 10, 1, 11)= 6.9833	 : orbe_theme( 10, 1, 12)= 5.9	 : orbe_theme( 10, 1, 13)= 3.0833	 : orbe_theme( 10, 1, 14)= 2.0666	 : orbe_theme( 10, 1, 15)= 0.4666
orbe_theme( 10, 2, 0)= 2.65	: orbe_theme( 10, 2, 1)= 0.1166	: orbe_theme( 10, 2, 2)= 0.5333	: orbe_theme( 10, 2, 3)= 0.8	: orbe_theme( 10, 2, 4)= 1.5166	: orbe_theme( 10, 2, 5)= 1.8	: orbe_theme( 10, 2, 6)= 0.6666	: orbe_theme( 10, 2, 7)= 0.7333
orbe_theme( 10, 2, 8)= 2.65	 : orbe_theme( 10, 2, 9)= 0.7333	 : orbe_theme( 10, 2, 10)= 0.6666	 : orbe_theme( 10, 2, 11)= 1.8	 : orbe_theme( 10, 2, 12)= 1.5166	 : orbe_theme( 10, 2, 13)= 0.8	 : orbe_theme( 10, 2, 14)= 0.5333	 : orbe_theme( 10, 2, 15)= 0.1166
orbe_theme( 10, 3, 0)= 6.3666	: orbe_theme( 10, 3, 1)= 0.2833	: orbe_theme( 10, 3, 2)= 1.2833	: orbe_theme( 10, 3, 3)= 1.9166	: orbe_theme( 10, 3, 4)= 3.65	: orbe_theme( 10, 3, 5)= 4.3166	: orbe_theme( 10, 3, 6)= 1.6166	: orbe_theme( 10, 3, 7)= 1.75
orbe_theme( 10, 3, 8)= 6.3666	 : orbe_theme( 10, 3, 9)= 1.75	 : orbe_theme( 10, 3, 10)= 1.6166	 : orbe_theme( 10, 3, 11)= 4.3166	 : orbe_theme( 10, 3, 12)= 3.65	 : orbe_theme( 10, 3, 13)= 1.9166	 : orbe_theme( 10, 3, 14)= 1.2833	 : orbe_theme( 10, 3, 15)= 0.2833
orbe_theme( 10, 4, 0)= 5.4166	: orbe_theme( 10, 4, 1)= 0.25	: orbe_theme( 10, 4, 2)= 1.0833	: orbe_theme( 10, 4, 3)= 1.6166	: orbe_theme( 10, 4, 4)= 3.1	: orbe_theme( 10, 4, 5)= 3.6666	: orbe_theme( 10, 4, 6)= 1.3666	: orbe_theme( 10, 4, 7)= 1.5
orbe_theme( 10, 4, 8)= 5.4166	 : orbe_theme( 10, 4, 9)= 1.5	 : orbe_theme( 10, 4, 10)= 1.3666	 : orbe_theme( 10, 4, 11)= 3.6666	 : orbe_theme( 10, 4, 12)= 3.1	 : orbe_theme( 10, 4, 13)= 1.6166	 : orbe_theme( 10, 4, 14)= 1.0833	 : orbe_theme( 10, 4, 15)= 0.25
orbe_theme( 10, 5, 0)= 6.45	: orbe_theme( 10, 5, 1)= 0.3	: orbe_theme( 10, 5, 2)= 1.3	: orbe_theme( 10, 5, 3)= 1.9333	: orbe_theme( 10, 5, 4)= 3.7	: orbe_theme( 10, 5, 5)= 4.3833	: orbe_theme( 10, 5, 6)= 1.6333	: orbe_theme( 10, 5, 7)= 1.7833
orbe_theme( 10, 5, 8)= 6.45	 : orbe_theme( 10, 5, 9)= 1.7833	 : orbe_theme( 10, 5, 10)= 1.6333	 : orbe_theme( 10, 5, 11)= 4.3833	 : orbe_theme( 10, 5, 12)= 3.7	 : orbe_theme( 10, 5, 13)= 1.9333	 : orbe_theme( 10, 5, 14)= 1.3	 : orbe_theme( 10, 5, 15)= 0.3
orbe_theme( 10, 6, 0)= 5.45	: orbe_theme( 10, 6, 1)= 0.25	: orbe_theme( 10, 6, 2)= 1.1	: orbe_theme( 10, 6, 3)= 1.6333	: orbe_theme( 10, 6, 4)= 3.1333	: orbe_theme( 10, 6, 5)= 3.7	: orbe_theme( 10, 6, 6)= 1.3833	: orbe_theme( 10, 6, 7)= 1.5
orbe_theme( 10, 6, 8)= 5.45	 : orbe_theme( 10, 6, 9)= 1.5	 : orbe_theme( 10, 6, 10)= 1.3833	 : orbe_theme( 10, 6, 11)= 3.7	 : orbe_theme( 10, 6, 12)= 3.1333	 : orbe_theme( 10, 6, 13)= 1.6333	 : orbe_theme( 10, 6, 14)= 1.1	 : orbe_theme( 10, 6, 15)= 0.25
orbe_theme( 10, 7, 0)= 3.9833	: orbe_theme( 10, 7, 1)= 0.1833	: orbe_theme( 10, 7, 2)= 0.8	: orbe_theme( 10, 7, 3)= 1.2	: orbe_theme( 10, 7, 4)= 2.2833	: orbe_theme( 10, 7, 5)= 2.7	: orbe_theme( 10, 7, 6)= 1	: orbe_theme( 10, 7, 7)= 1.1
orbe_theme( 10, 7, 8)= 3.9833	 : orbe_theme( 10, 7, 9)= 1.1	 : orbe_theme( 10, 7, 10)= 1	 : orbe_theme( 10, 7, 11)= 2.7	 : orbe_theme( 10, 7, 12)= 2.2833	 : orbe_theme( 10, 7, 13)= 1.2	 : orbe_theme( 10, 7, 14)= 0.8	 : orbe_theme( 10, 7, 15)= 0.1833
orbe_theme( 10, 8, 0)= 3.45	: orbe_theme( 10, 8, 1)= 0.15	: orbe_theme( 10, 8, 2)= 0.6833	: orbe_theme( 10, 8, 3)= 1.0333	: orbe_theme( 10, 8, 4)= 1.9833	: orbe_theme( 10, 8, 5)= 2.3333	: orbe_theme( 10, 8, 6)= 0.8666	: orbe_theme( 10, 8, 7)= 0.95
orbe_theme( 10, 8, 8)= 3.45	 : orbe_theme( 10, 8, 9)= 0.95	 : orbe_theme( 10, 8, 10)= 0.8666	 : orbe_theme( 10, 8, 11)= 2.3333	 : orbe_theme( 10, 8, 12)= 1.9833	 : orbe_theme( 10, 8, 13)= 1.0333	 : orbe_theme( 10, 8, 14)= 0.6833	 : orbe_theme( 10, 8, 15)= 0.15
orbe_theme( 10, 9, 0)= 3.3333	: orbe_theme( 10, 9, 1)= 0.15	: orbe_theme( 10, 9, 2)= 0.6666	: orbe_theme( 10, 9, 3)= 1	: orbe_theme( 10, 9, 4)= 1.9166	: orbe_theme( 10, 9, 5)= 2.2666	: orbe_theme( 10, 9, 6)= 0.85	: orbe_theme( 10, 9, 7)= 0.9166
orbe_theme( 10, 9, 8)= 3.3333	 : orbe_theme( 10, 9, 9)= 0.9166	 : orbe_theme( 10, 9, 10)= 0.85	 : orbe_theme( 10, 9, 11)= 2.2666	 : orbe_theme( 10, 9, 12)= 1.9166	 : orbe_theme( 10, 9, 13)= 1	 : orbe_theme( 10, 9, 14)= 0.6666	 : orbe_theme( 10, 9, 15)= 0.15
orbe_theme( 10, 10, 0)= 0	: orbe_theme( 10, 10, 1)= 0	: orbe_theme( 10, 10, 2)= 0	: orbe_theme( 10, 10, 3)= 0	: orbe_theme( 10, 10, 4)= 0	: orbe_theme( 10, 10, 5)= 0	: orbe_theme( 10, 10, 6)= 0	: orbe_theme( 10, 10, 7)= 0
orbe_theme( 10, 10, 8)= 0	 : orbe_theme( 10, 10, 9)= 0	 : orbe_theme( 10, 10, 10)= 0	 : orbe_theme( 10, 10, 11)= 0	 : orbe_theme( 10, 10, 12)= 0	 : orbe_theme( 10, 10, 13)= 0	 : orbe_theme( 10, 10, 14)= 0	 : orbe_theme( 10, 10, 15)= 0
orbe_theme( 10, 11, 0)= 0	: orbe_theme( 10, 11, 1)= 0	: orbe_theme( 10, 11, 2)= 0	: orbe_theme( 10, 11, 3)= 0	: orbe_theme( 10, 11, 4)= 0	: orbe_theme( 10, 11, 5)= 0	: orbe_theme( 10, 11, 6)= 0	: orbe_theme( 10, 11, 7)= 0
orbe_theme( 10, 11, 8)= 0	 : orbe_theme( 10, 11, 9)= 0	 : orbe_theme( 10, 11, 10)= 0	 : orbe_theme( 10, 11, 11)= 0	 : orbe_theme( 10, 11, 12)= 0	 : orbe_theme( 10, 11, 13)= 0	 : orbe_theme( 10, 11, 14)= 0	 : orbe_theme( 10, 11, 15)= 0


'Lilith, Lune à Lilith (0 à 11), aspects (0 à 15)
orbe_theme( 11, 0, 0)= 10.3166	: orbe_theme( 11, 0, 1)= 0.4666	: orbe_theme( 11, 0, 2)= 2.0666	: orbe_theme( 11, 0, 3)= 3.1	: orbe_theme( 11, 0, 4)= 5.9166	: orbe_theme( 11, 0, 5)= 7	: orbe_theme( 11, 0, 6)= 2.6166	: orbe_theme( 11, 0, 7)= 2.85
orbe_theme( 11, 0, 8)= 10.3166	 : orbe_theme( 11, 0, 9)= 2.85	 : orbe_theme( 11, 0, 10)= 2.6166	 : orbe_theme( 11, 0, 11)= 7	 : orbe_theme( 11, 0, 12)= 5.9166	 : orbe_theme( 11, 0, 13)= 3.1	 : orbe_theme( 11, 0, 14)= 2.0666	 : orbe_theme( 11, 0, 15)= 0.4666
orbe_theme( 11, 1, 0)= 10.2833	: orbe_theme( 11, 1, 1)= 0.4666	: orbe_theme( 11, 1, 2)= 2.0666	: orbe_theme( 11, 1, 3)= 3.0833	: orbe_theme( 11, 1, 4)= 5.9	: orbe_theme( 11, 1, 5)= 6.9833	: orbe_theme( 11, 1, 6)= 2.6166	: orbe_theme( 11, 1, 7)= 2.85
orbe_theme( 11, 1, 8)= 10.2833	 : orbe_theme( 11, 1, 9)= 2.85	 : orbe_theme( 11, 1, 10)= 2.6166	 : orbe_theme( 11, 1, 11)= 6.9833	 : orbe_theme( 11, 1, 12)= 5.9	 : orbe_theme( 11, 1, 13)= 3.0833	 : orbe_theme( 11, 1, 14)= 2.0666	 : orbe_theme( 11, 1, 15)= 0.4666
orbe_theme( 11, 2, 0)= 2.65	: orbe_theme( 11, 2, 1)= 0.1166	: orbe_theme( 11, 2, 2)= 0.5333	: orbe_theme( 11, 2, 3)= 0.8	: orbe_theme( 11, 2, 4)= 1.5166	: orbe_theme( 11, 2, 5)= 1.8	: orbe_theme( 11, 2, 6)= 0.6666	: orbe_theme( 11, 2, 7)= 0.7333
orbe_theme( 11, 2, 8)= 2.65	 : orbe_theme( 11, 2, 9)= 0.7333	 : orbe_theme( 11, 2, 10)= 0.6666	 : orbe_theme( 11, 2, 11)= 1.8	 : orbe_theme( 11, 2, 12)= 1.5166	 : orbe_theme( 11, 2, 13)= 0.8	 : orbe_theme( 11, 2, 14)= 0.5333	 : orbe_theme( 11, 2, 15)= 0.1166
orbe_theme( 11, 3, 0)= 6.3666	: orbe_theme( 11, 3, 1)= 0.2833	: orbe_theme( 11, 3, 2)= 1.2833	: orbe_theme( 11, 3, 3)= 1.9166	: orbe_theme( 11, 3, 4)= 3.65	: orbe_theme( 11, 3, 5)= 4.3166	: orbe_theme( 11, 3, 6)= 1.6166	: orbe_theme( 11, 3, 7)= 1.75
orbe_theme( 11, 3, 8)= 6.3666	 : orbe_theme( 11, 3, 9)= 1.75	 : orbe_theme( 11, 3, 10)= 1.6166	 : orbe_theme( 11, 3, 11)= 4.3166	 : orbe_theme( 11, 3, 12)= 3.65	 : orbe_theme( 11, 3, 13)= 1.9166	 : orbe_theme( 11, 3, 14)= 1.2833	 : orbe_theme( 11, 3, 15)= 0.2833
orbe_theme( 11, 4, 0)= 5.4166	: orbe_theme( 11, 4, 1)= 0.25	: orbe_theme( 11, 4, 2)= 1.0833	: orbe_theme( 11, 4, 3)= 1.6166	: orbe_theme( 11, 4, 4)= 3.1	: orbe_theme( 11, 4, 5)= 3.6666	: orbe_theme( 11, 4, 6)= 1.3666	: orbe_theme( 11, 4, 7)= 1.5
orbe_theme( 11, 4, 8)= 5.4166	 : orbe_theme( 11, 4, 9)= 1.5	 : orbe_theme( 11, 4, 10)= 1.3666	 : orbe_theme( 11, 4, 11)= 3.6666	 : orbe_theme( 11, 4, 12)= 3.1	 : orbe_theme( 11, 4, 13)= 1.6166	 : orbe_theme( 11, 4, 14)= 1.0833	 : orbe_theme( 11, 4, 15)= 0.25
orbe_theme( 11, 5, 0)= 6.45	: orbe_theme( 11, 5, 1)= 0.3	: orbe_theme( 11, 5, 2)= 1.3	: orbe_theme( 11, 5, 3)= 1.9333	: orbe_theme( 11, 5, 4)= 3.7	: orbe_theme( 11, 5, 5)= 4.3833	: orbe_theme( 11, 5, 6)= 1.6333	: orbe_theme( 11, 5, 7)= 1.7833
orbe_theme( 11, 5, 8)= 6.45	 : orbe_theme( 11, 5, 9)= 1.7833	 : orbe_theme( 11, 5, 10)= 1.6333	 : orbe_theme( 11, 5, 11)= 4.3833	 : orbe_theme( 11, 5, 12)= 3.7	 : orbe_theme( 11, 5, 13)= 1.9333	 : orbe_theme( 11, 5, 14)= 1.3	 : orbe_theme( 11, 5, 15)= 0.3
orbe_theme( 11, 6, 0)= 5.45	: orbe_theme( 11, 6, 1)= 0.25	: orbe_theme( 11, 6, 2)= 1.1	: orbe_theme( 11, 6, 3)= 1.6333	: orbe_theme( 11, 6, 4)= 3.1333	: orbe_theme( 11, 6, 5)= 3.7	: orbe_theme( 11, 6, 6)= 1.3833	: orbe_theme( 11, 6, 7)= 1.5
orbe_theme( 11, 6, 8)= 5.45	 : orbe_theme( 11, 6, 9)= 1.5	 : orbe_theme( 11, 6, 10)= 1.3833	 : orbe_theme( 11, 6, 11)= 3.7	 : orbe_theme( 11, 6, 12)= 3.1333	 : orbe_theme( 11, 6, 13)= 1.6333	 : orbe_theme( 11, 6, 14)= 1.1	 : orbe_theme( 11, 6, 15)= 0.25
orbe_theme( 11, 7, 0)= 3.9833	: orbe_theme( 11, 7, 1)= 0.1833	: orbe_theme( 11, 7, 2)= 0.8	: orbe_theme( 11, 7, 3)= 1.2	: orbe_theme( 11, 7, 4)= 2.2833	: orbe_theme( 11, 7, 5)= 2.7	: orbe_theme( 11, 7, 6)= 1	: orbe_theme( 11, 7, 7)= 1.1
orbe_theme( 11, 7, 8)= 3.9833	 : orbe_theme( 11, 7, 9)= 1.1	 : orbe_theme( 11, 7, 10)= 1	 : orbe_theme( 11, 7, 11)= 2.7	 : orbe_theme( 11, 7, 12)= 2.2833	 : orbe_theme( 11, 7, 13)= 1.2	 : orbe_theme( 11, 7, 14)= 0.8	 : orbe_theme( 11, 7, 15)= 0.1833
orbe_theme( 11, 8, 0)= 3.45	: orbe_theme( 11, 8, 1)= 0.15	: orbe_theme( 11, 8, 2)= 0.6833	: orbe_theme( 11, 8, 3)= 1.0333	: orbe_theme( 11, 8, 4)= 1.9833	: orbe_theme( 11, 8, 5)= 2.3333	: orbe_theme( 11, 8, 6)= 0.8666	: orbe_theme( 11, 8, 7)= 0.95
orbe_theme( 11, 8, 8)= 3.45	 : orbe_theme( 11, 8, 9)= 0.95	 : orbe_theme( 11, 8, 10)= 0.8666	 : orbe_theme( 11, 8, 11)= 2.3333	 : orbe_theme( 11, 8, 12)= 1.9833	 : orbe_theme( 11, 8, 13)= 1.0333	 : orbe_theme( 11, 8, 14)= 0.6833	 : orbe_theme( 11, 8, 15)= 0.15
orbe_theme( 11, 9, 0)= 3.3333	: orbe_theme( 11, 9, 1)= 0.15	: orbe_theme( 11, 9, 2)= 0.6666	: orbe_theme( 11, 9, 3)= 1	: orbe_theme( 11, 9, 4)= 1.9166	: orbe_theme( 11, 9, 5)= 2.2666	: orbe_theme( 11, 9, 6)= 0.85	: orbe_theme( 11, 9, 7)= 0.9166
orbe_theme( 11, 9, 8)= 3.3333	 : orbe_theme( 11, 9, 9)= 0.9166	 : orbe_theme( 11, 9, 10)= 0.85	 : orbe_theme( 11, 9, 11)= 2.2666	 : orbe_theme( 11, 9, 12)= 1.9166	 : orbe_theme( 11, 9, 13)= 1	 : orbe_theme( 11, 9, 14)= 0.6666	 : orbe_theme( 11, 9, 15)= 0.15
orbe_theme( 11, 10, 0)= 0	: orbe_theme( 11, 10, 1)= 0	: orbe_theme( 11, 10, 2)= 0	: orbe_theme( 11, 10, 3)= 0	: orbe_theme( 11, 10, 4)= 0	: orbe_theme( 11, 10, 5)= 0	: orbe_theme( 11, 10, 6)= 0	: orbe_theme( 11, 10, 7)= 0
orbe_theme( 11, 10, 8)= 0	 : orbe_theme( 11, 10, 9)= 0	 : orbe_theme( 11, 10, 10)= 0	 : orbe_theme( 11, 10, 11)= 0	 : orbe_theme( 11, 10, 12)= 0	 : orbe_theme( 11, 10, 13)= 0	 : orbe_theme( 11, 10, 14)= 0	 : orbe_theme( 11, 10, 15)= 0
orbe_theme( 11, 11, 0)= 0	: orbe_theme( 11, 11, 1)= 0	: orbe_theme( 11, 11, 2)= 0	: orbe_theme( 11, 11, 3)= 0	: orbe_theme( 11, 11, 4)= 0	: orbe_theme( 11, 11, 5)= 0	: orbe_theme( 11, 11, 6)= 0	: orbe_theme( 11, 11, 7)= 0
orbe_theme( 11, 11, 8)= 0	 : orbe_theme( 11, 11, 9)= 0	 : orbe_theme( 11, 11, 10)= 0	 : orbe_theme( 11, 11, 11)= 0	 : orbe_theme( 11, 11, 12)= 0	 : orbe_theme( 11, 11, 13)= 0	 : orbe_theme( 11, 11, 14)= 0	 : orbe_theme( 11, 11, 15)= 0

end sub

sub temp_orbes
dim sujet, agent, aspect as integer
dim orbemax as double
dim car, car1, car2 as double
dim rang, col as integer
dim car3 as currency

Doc = ThisComponent
Sheet = Doc.Sheets.getByName("temp")
sheet.charheight=6
'goto debut
'conversion minutes en centiemes : ne faire qu'une fois !
for i=0 to 107
cell=Sheet.getCellByPosition(0, i)
abc=cell.getstring
car=instr(1,abc,",")
	if car then
	car1=val(mid(abc,1,car-1))*60
	car2=val(mid(abc,car+1))
	car3=car1+car2
	car3=car3/60
	else
	car3=val(abc)
	endif
cell.value=car3
Cell.NumberFormat =10002 '2 décimales affichées
next i

debut:
sujet=11
agent=0
aspect=0
	'effacement
	Cell = Sheet.getCellrangeByPosition(1,0,20,200) 
	cell.clearcontents (1 or 2 or 4 or 32)
	sheet.charheight=6
for i=0 to 107 step 9
		for j=0 to 8
		orbemax=Sheet.getCellByPosition(0, i+j).getvalue
		 Sheet.getCellByPosition(1, i+j).string= "orbe_theme(" & str(sujet) & "," & str(agent) & "," & str(aspect) & ")" &  "=" & str(orbemax)
		 aspect=aspect+1
	 	next j
		'calcul et regroupement des aspects descendants
		rang=i+j-1
		col=2
		aspect=9
		do until aspect=16
		orbemax=Sheet.getCellByPosition(0, rang-1).getvalue
		Sheet.getCellByPosition(col, i+j-1).string= " : orbe_theme(" & str(sujet) & "," & str(agent) & "," & str(aspect) & ")" &  "=" & str(orbemax)
		col=col+1
		rang=rang-1
		aspect=aspect+1
		loop
		'regroupement des aspects croissants
		rang=i+j-3
		col=2
		for k=i to rang
		abc=Sheet.getCellByPosition(1, k+1).getstring
		Sheet.getCellByPosition(col, i).string= ": " & abc
		col=col+1
		next k
aspect=0
agent=agent+1
next i
	'compression
	for i=1 to 107 step 2
	CellRangeAddress.Sheet = 1
	CellRangeAddress.StartColumn = 1
	CellRangeAddress.StartRow = i
	CellRangeAddress.EndColumn = 20
	CellRangeAddress.EndRow = i+6'i+1
	Sheet.removeRange(CellRangeAddress, com.sun.star.sheet.CellDeleteMode.UP)
	next i
	
	'ajustement de la largeur des colonnes
	Sheet = Doc.Sheets.getByName("temp")
	for i=0 to 33
	Sheet.columns(i).Optimalwidth = True
	next i
end sub


sub totale
calcul_du_theme
ephemerides
phases_ephemerides
calcul_des_aspects
tableau_transits
tableau_par_planete
tableau_par_annee
end sub


Sub a4_graphe_du_jour
Dim Charts As Object
Dim Chart as Object
Dim Rect As New com.sun.star.awt.Rectangle
Dim RangeAddress(0 to 1) As New com.sun.star.table.CellRangeAddress

dim jours as long
dim position(1 to 12)as double
dim position0, position1, position2 as double
dim orbe as double
dim orbeentier as integer
dim signemoins as string
Dim CellAddress As New com.sun.star.table.CellAddress

Doc = ThisComponent
'vérifie si feuille éphémérides présnte
If not Doc.Sheets.hasByName("éphémérides") Then msgbox "pas de feuille éphémérides !, abandon" : exit sub

'sélectionne 1ère feuille, cellule A16 (éphémérides)
Sheet = Doc.Sheets.getByName("éphémérides")
if Sheet.getCellByPosition(1, 1).getstring="" then  msgbox "le thème est vide, exécuter 'calcul_du_thème' d'abord" : exit sub
if Sheet.getCellByPosition(0, 17).getstring="" then  msgbox "le tableau des éphémérides est vide, exécuter 'ephemerides' d'abord" : exit sub
'suppression et recréation feuille graphe 
If Doc.Sheets.hasByName("graphe") Then  Doc.Sheets.RemoveByName("graphe")'suppression feuille
Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
Doc.Sheets.insertByName("graphe", Sheet)


'exemple pour déterminer le type de contenu de la  cellule A16 (= "text")
'Select Case Cell.Type 
'Case com.sun.star.table.CellContentType.EMPTY 
 ' MsgBox "Content: Empty"
'Case com.sun.star.table.CellContentType.VALUE
  ' MsgBox "Content: Value"
'Case com.sun.star.table.CellContentType.TEXT
 ' MsgBox "Content: Text"
'Case com.sun.star.table.CellContentType.FORMULA
  ' MsgBox "Content: Formula"
'End Select

' lit l'orbe max
	Sheet = Doc.Sheets.getByName("éphémérides")
	orbe= Sheet.getCellByPosition(0, 0).getvalue

'récupère le nombre de lignes utilisées dans la feuille 1
	Sheet = Doc.Sheets.getByName("éphémérides")
	Curs = Sheet.createCursor
	Curs.gotoEndOfUsedArea(True)
	lastrow = Curs.Rows.Count -1
 
'calcule le nombre de jours entre date du jour et le dernier jour des éphémérides
Cell = Sheet.getCellByPosition(0, lastrow)
abc=cell.getstring
jours=datevalue(date)-datevalue(cell.getstring) 'cell.getvalue si nombre ?
if jours > 0 then msgbox "date du jour postérieure à la date de fin des éphémérides, pas de graphique du jour" : goto fin
'calcule le nombre de jours entre date du jour et le 1er jour des éphémérides
Cell = Sheet.getCellByPosition(0, 17)
jours=datevalue(date)-datevalue(cell.getstring) 'cell.getvalue si nombre ?
if jours < 0 then msgbox "date du jour antérieure à la date de départ des éphémérides, pas de graphique du jour" : goto fin
'récupération des positions des planètes du jour
for i = 1 to 10
position(i)= Sheet.getCellByPosition(i, 17+jours).getvalue
if position(i)=0 then msgbox "ligne " & str(jours+17) & " vide ou incomplète, pas de graphique du jour" : goto fin
next i

' conversion orbe maxi en degrés et minutes
coeff1=int(orbe/30)*30 : coeff2=int(orbe) mod 30
bcd= str(coeff2) & chr$(176) & str(int(60*(orbe-coeff1-coeff2))) & "')"

abc=Sheet.getCellByPosition(19, 0).getstring
abc=abc & " - Transits du " & date & " (orbe maxi" & bcd

'compare données du jour à données du theme (abc= liste des transits)
for i=1 to 16' aspect
for j= 1 to 16 '10 'planète du thème
position0 = Sheet.getCellByPosition(j, i).getvalue
position1 = position0-orbe : position2 = position0 + orbe 'tolérance = +- 0,5
for k=1 to 12 'planète du jour
	if position(k) > position1 and position(k) < position2 then
	abc=abc & chr$(13) & Sheet.getCellByPosition(k, 0).getstring & " " 'planete en transit
	abc=abc & Sheet.getCellByPosition(0, i).getstring & "  " 'aspect
	abc=abc &  Sheet.getCellByPosition(j, 0).getstring  'planete transitée
	
	orbedecimal= position(k) - position0 : orbeentier=orbedecimal
	signemoins="+"
	if orbedecimal < 0 then orbedecimal=abs(orbedecimal) : signemoins="-"
	coeff1=int(orbedecimal/30)*30 : coeff2=int(orbedecimal) mod 30
	bcd=signemoins & str(coeff2) & chr$(176) & str(int(60*(orbedecimal-coeff1-coeff2))) & "'"' & "' " &signe(coeff1/30) 'ex  12° 23' Lion
		
	abc=abc & " (orbe " & bcd &")"'orbe
	endif
next k
next j
next i

'!! attention vérifier qu'il n'y a pas d'autre feuille avec un graphe sinon erreur !!
'affichage graphe 
'sélectionne 3ème feuille
'Sheet = Doc.Sheets (numero_feuille)
Sheet = Doc.Sheets.getByName("graphe")
'charts=Doc.Sheets(numero_feuille).Charts
charts=sheet.Charts
'efface graphe
Charts.removebyname("MyChart")
Rect.X = 800
Rect.Y = 100
Rect.Width = 19000
Rect.Height = 9000
'sélection cellules theme
RangeAddress(0).Sheet = 0
RangeAddress(0).StartColumn =0
RangeAddress(0).StartRow =0
RangeAddress(0).EndColumn = 12
RangeAddress(0).EndRow =14
'sélection ligne du jour
RangeAddress(1).Sheet = 0
RangeAddress(1).StartColumn =0
RangeAddress(1).StartRow =17+jours
RangeAddress(1).EndColumn = 12
RangeAddress(1).EndRow =17+jours
'tracé graphe à barres
Charts.addNewByName("MyChart", Rect, RangeAddress(), True, True)
'changement du type de graphe
Chart = Charts.getByName("MyChart").embeddedObject
Chart.Diagram = Chart.createInstance("com.sun.star.chart.LineDiagram")
Chart.HasLegend = True
Chart.legend.CharHeight = 7
chart.diagram.symboltype=9'auto
chart.diagram.XAxis.charheight=7
chart.diagram.YAxis.charheight=7
chart.diagram.HasXAxisGrid=true
Chart.Diagram.Lines =false
chart.title.string=abc 'liste des transits du jour
Chart.title.CharHeight = 7

fin:
'mise au 1er plan de la feuille
Sheet = Doc.Sheets.getByName("graphe")
Controller = Doc.CurrentController
controller.setActiveSheet(sheet)

End sub

REM  *****  BASIC  *****


'1954
Sub calcul_du_theme

dim choixnom as integer
dim compteur as integer
dim nomfichier as string
dim nommax as integer
dim pos1 as double
dim nom(1 to 100) as string
dim position(1 to 16) as double 'currency
dim bb

goto debut0

'obsolète - récupère la liste des fichiers .txt
	nommax=0
	abc=Dir(curdir & "/*.txt")
	if abc = "" then
	bcd=inputbox ("pas de fichier '.txt' trouvé ! utilisation du thème de la 2ème feuille ?", "lecture thème","O")
	if bcd ="" then exit sub else choix=2 : goto debut
	else
	car=1
		do until abc=""
		abc=dir()
		if abc <> "" and instr (1,abc,"ephe") =0 then nom(car)=abc : car=car+1
		loop
	endif
	nommax=car-1


debut0:
'option fichier ou thème interne
choix=inputbox ("lecture thème par fichier .txt (1) ou d'après le contenu de la feuille thème (2) ?", "choix lecture thème","1")
if val(choix) <=0  or val(choix) > 2 then exit sub


debut:
Doc = ThisComponent

select case val(choix)

'****************écrture valeurs de la conjonction à partir d'un fichier .txt***************
case 1
 'laisser dans le fichier le séparateur décimal = "." (point)
	'ouvre sélecteur de fichiers
	abc=""
	call FileButtonSelected
	if abc="" then exit sub
	nomfichier=abc 
	goto debut1

'obsolète - liste des fichiers .txt
	abc=""
	for i=1 to nommax
	abc=abc & str(i) & "=" &  nom(i) &","'& chr$(13)
	next i
' obsolète - choix du fichier .txt
	question:
	choixnom=inputbox (abc, "choix fichier","1")
	if choixnom=0 then exit sub
	if choixnom > nommax or choixnom < 0  then goto question
	nomfichier=nom(choixnom)

debut1:
'confirmation du nom de fichier
	bcd=inputbox ("confirmer choix du thème : " & nomfichier, nomfichier,"O")
	'if bcd <> "O" then goto question
	if bcd <> "O" then exit sub

'désactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(false)
	
'initialisation feuille éphémérides
call feuille_ephemerides


'ouverture fichier .txt
on error goto fin
	open nomfichier for input as #1

Do While Not EOF(1)
   Line Input #1, abc
   'vérifie si ligne des coordonnées de naissance
  	car1=instr(1,abc,"GMT") 
	   if car1 then
		   'écriture des coordonnées colonne W
		   		Sheet.getCellByPosition(22, 0).string=abc 
		   'écriture date de naissance colonne W
		   		bcd=mid$(abc,5,10) 'jj-mm-aaaa
			   'suppression des espaces
			   bcd= Replace$(bcd," ", "") 
			   date_naissance=datevalue(bcd)
			   cell=Sheet.getCellByPosition(22, 1)
			   cell.value=date_naissance
			   'format dd/mm/aaaa
		   		cell.numberformat=10030
		   'heure de naissance
				Sheet.getCellByPosition(22, 2).string=mid$(abc,16,5)
			'copie date et heure de naissance feuille thème (si existe), colonne F
			If Doc.Sheets.hasByName("thème") then
		   		sheet=Doc.Sheets.getByName("thème")
		   		cell=sheet.getCellByPosition(5, 13)
		   		cell.value=date_naissance
		   		cell.numberformat=10030
				sheet.getCellByPosition(6, 13).string=mid$(abc,16,5)
				Sheet = Doc.Sheets.getByName("éphémérides")
			endif
	   endif
	   
   'vérifie si ligne position Soleil avec Zet9 (en anglais)
   car1=instr(1,abc,"Sun") 
  	 'vérifie si ligne position Soleil avec Astrolog32 (en français)
  	 if car1=0 then car1=instr(1,abc,"Soleil") 
  	 if car1 then
      	'lit les positions de Soleil à Pluton (car = 1 à 10) + Noeud Nord, AS, FC, DS et MC (car = 11 à 15)
      	car=1 'nombre de positions trouvées
      	Do While Not EOF(1)
      	car2=instr(3,abc,".") 'position du point décimal de la longitude
	   		if car2 then
	   		'position planète
	   		bcd = mid(abc,car2 -3,7)' longitude
	   		Cell = Sheet.getCellByPosition(car, 1) : cell.value=val(bcd) 'écriture 1ère ligne de données des aspects (conjonction)
		   		'rétrograde ?
		   		if mid(abc,car2+5,1) ="R" then
		   		Cell = Sheet.getCellByPosition(car,0) :cell.string=cell.string & " (R)"
	   			endif
	   		'position cuspide maison
	   		car3=instr(1,abc,"cuspide")
		   		if car3 then
		   		car2=instr(car3,abc,".")'position du point décimal de la position de la cuspide
			   		if car2 then
			   		bcd = mid(abc,car2 -3,7)' longitude
			   		Cell = Sheet.getCellByPosition(21, car) : cell.value=val(bcd) 'écriture colonne 21
			   		endif
			   	endif
		   	car=car+1
	   		endif
   		Line Input #1, abc
   		loop
  	  	exit do
  	 endif
Loop
Close #1

'vérifie si données complètes
if car1=0 then  msgbox "fichier incorrect, pas de position trouvée pour le Soleil" : exit sub
if car < 17 then msgbox "données incomplètes, abandon" : exit sub

'écriture du nom de fichier
	'suppression du chemin dans le nom de fichier
	car1=1
	do
	car2=instr(car1,nomfichier,"/")
	if car2=0 then exit do
	car1=car2+1
	loop
	nomfichier=mid$(nomfichier,car1)
	'supprime l'extension .txt
	Cell = Sheet.getCellByPosition(19, 0)
	cell.string=mid$(nomfichier,1,instr(1,nomfichier,".txt")-1) '"Astro.txt"
	Cell.CellBackColor = RGB(0, 255, 0)


'***********écrture valeurs de la conjonction à partir des données du thème de la page thème************
case 2
'confirmation
	choix=inputbox ("confirmer calcul du thème (O,N) ?", "calcul du thème","O")
	if choix <> "O" then exit sub

'désactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(false)

'création feuille thème  si absente
	If not Doc.Sheets.hasByName("thème") Then call feuille_theme

'lecture données du thème
Sheet = Doc.Sheets.getByName("thème")
	'Soleil à MC
	for i= 1 to 16
		'signes
		for j = 1 to 12
			'degrés
			abc=Sheet.getCellByPosition(i, j).getstring
			if abc <> "" then
				'recherche position du symbole "degrés"
				car1=instr(1,abc,chr$(176))
				'sinon recherche du séparateur "."
				if car1=0 then car1=instr(1,abc,".")
				if car1=0 then msgbox "manque symbole '.' pour la séparation degrés/minutes de " & planete(i-1) : exit sub
				degres= val(mid(abc,1,car1-1))
				'minutes
				minutes= val(mid(abc,car1+1))
				if degres > 30 or minutes > 59 then msgbox "abandon, valeur erronée pour la planète " & planete(i-1) : exit sub
				position(i)=(degres*60 + minutes)/60 +30*(j-1)
				exit for
			endif
		next j
	Next i

'vérification si thème complet
	for i=1 to 16
	if position(i)=0 then msgbox "abandon, données manquantes planète : " & planete(i-1) : exit sub
	if position(i) > 360 then msgbox "abandon, données incohérentes pour la planète : " & planete(i-1) : exit sub
	next i
	
'initialisation feuille éphémérides
	call feuille_ephemerides
	
'effacement ancien thème, feuille épéhmérides
	Sheet = Doc.Sheets.getByName("éphémérides")
	cell= sheet.getCellRangeByPosition(1,1,16,16)
	cell.clearcontents (1 or 2 or 4) '1 : valeurs numériques, 2 : date, 4 : string, 32: formatage (dont couleur et avec recalcul du lastrow)

'écriture ligne "conjonction" du thème sur la feuille éphémérides 
	for i=1 to 16
	Cell = Sheet.getCellByPosition(i, 1) : cell.value=position(i)
	next i

'cellule avec nom du thème + couleur
	Cell = Sheet.getCellByPosition(19, 0) : cell.string="feuille thème"
	Cell.CellBackColor = RGB(0, 255, 0)

'écrit date et heure de naissance dans feuille éphémérides, si présentes dans feuiile thème
	'lecture
		Sheet = Doc.Sheets.getByName("thème")
	'date de naissance (valeur numérique)
		date_naissance=sheet.getCellByPosition(5, 13).getvalue
	'heure de naissance (string)
		abc=sheet.getCellByPosition(6, 13).getstring
	'écriture
		Sheet = Doc.Sheets.getByName("éphémérides")
		cell=Sheet.getCellByPosition(22, 1)
		cell.value=date_naissance
	   'format dd/mm/aaaa
		cell.numberformat=10030
	'heure de naissance
		Sheet.getCellByPosition(22, 2).string=abc
end select


'*******************case 1 et case 2 : remplissage du tableau des aspects feuille éphémérides, à partir de la ligne conjonction************
Sheet = Doc.Sheets.getByName("éphémérides")
	for i = 1 to 16 'planète
		for j = 2 to 16 'aspect
			car3=30
			if j = 3 or j=4 or j=7 or j=8 or j=11 or j=12 or j=15 or j =16 then car3=15
			'si données ligne conjonction, écriture des autres aspects
			if Sheet.getCellByPosition(i, j-1).getstring <>"" then
			pos1= Sheet.getCellByPosition(i, j-1).getvalue + car3
			if pos1>360 then pos1 =pos1- 360
			Cell = Sheet.getCellByPosition(i, j) : cell.value=pos1
			else
			exit for
			endif
		next j
	next i

'thème à 2 décimales
'cell=Sheet.getCellrangeByPosition(1,1,16,16) 
'bb=oFA.callFunction("Round",array(cell,2))
'cell.setdata(bb)

'format orbes = 2 décimales affichées
	Cell = Sheet.getCellrangeByPosition(17,1,18,16) 
	cell.numberformat=10002
	

'pas de décimales affichées pour les données du thème
	Cell = Sheet.getCellrangeByPosition(1,1,16,16) 
	cell.numberformat=10001


'ajustement de la largeur des colonnes (prend du temps quand il y a des éphémérides !)
	for i=20 to 22 'au lieu de 0 to 22, trop long !
	Sheet.columns(i).Optimalwidth = True
	next i

'mise en couleur jaune clair
Cell = Sheet.getCellrangeByPosition(0,0,16,16) : cell.cellBackColor = RGB(255, 255, 155)


'mise au 1er plan de la feuille
	Sheet = Doc.Sheets.getByName("éphémérides")
	Controller = Doc.CurrentController
	controller.setActiveSheet(sheet)
	
'affichage du nouveau nom dans la form (si form 'visuel' active)
	if form_ok=1 then
	'nom
	commande1 = feuille.getControl("Label2")
	commande1.text=Sheet.getCellByPosition(19, 0).getstring
	'date de naissance
	commande1 = feuille.getControl("Label11")
	commande1.text=Sheet.getCellByPosition(22, 1).getstring
	endif
	
'active boutons tableaux, thème, aspects, etc.
call actions_boutons2(true)
		
'call zodiaque(1,1)
call calcul_des_aspects
	'msgbox "terminé"
exit sub

fin:
msgbox "erreur lecture fichier " & nomfichier : exit sub
End Sub


Sub feuille_ephemerides
'définitions variables aspect et planete du theme
call definitions

'création feuille éphémérides si absente
	If not Doc.Sheets.hasByName("éphémérides") Then 
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
	Doc.Sheets.insertByName("éphémérides", Sheet)
	endif

Sheet = Doc.Sheets.getByName("éphémérides")	
'taille caractères et définition orbe du graphe
	sheet.charheight=7
	'orbe utilisé par le graphe du jour
	Sheet.getCellByPosition(0,0).value=1

'en-têtes des colonnes feuille éphémérides (planètes)
	for i = 0 to 15
	Cell = Sheet.getCellByPosition(i+1,0) : cell.string=planete(i)
	Cell.HoriJustify = com.sun.star.table.CellVertJustify.CENTER
	next i
	Sheet.getCellByPosition(17,0).string="orbe thème"
	Sheet.getCellByPosition(18,0).string="orbe transit"
	Sheet.getCellByPosition(20,0).string="Maison"
	Sheet.getCellByPosition(21,0).string="longitude"
	
'effacement positions Maisons colonnes U,V,W
	cell= sheet.getCellRangeByPosition(20,1,22,16)
	cell.clearcontents (1 or 2 or 4) 

'effacement nom du thème et coordonnées de naissance
	Sheet.getCellByPosition(19,0).string=""
	Sheet.getCellByPosition(22,0).string=""
	
'en-têtes des lignes (aspects) + orbes max colonnes Q et R 
	for i =0 to 15
	Sheet.getCellByPosition(0, i+1).string=aspect(i) 'aspect
	Sheet.getCellByPosition(17,i+1).value=orbe(i)  'orbe max du thème
	Sheet.getCellByPosition(18,i+1).value=orbe_transit(i)  'orbe max du transit
	if i <=11 then Sheet.getCellByPosition(20,i+1).value= i+1
	next i
	
'effacement données du thème
	cell= sheet.getCellRangeByPosition(1,1,16,16)
	cell.clearcontents (1 or 2 or 4) '1 : valeurs numériques, 2 : date, 4 : string, 32: formatage (dont couleur et avec recalcul du lastrow)
	
End Sub




Sub feuille_theme
Dim aBorder as New com.sun.star.table.BorderLine
Dim oBorder as New com.sun.star.table.TableBorder
dim yy() 'pour lignes deg/min et colonnes signes
dim zz(0,15) 'pour écrire lignes planètes

Doc = ThisComponent
Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
Doc.Sheets.insertByName("thème", Sheet)

'définitions variables aspect et planete du theme
call definitions

Sheet = Doc.Sheets.getByName("thème")

'***partie supérieure (thème)***

'mise en couleur d'une ligne sur deux dans le tableau du thème
	for i=0 to 12 step 2
		sheet.getCellRangeByPosition(0,i,17,i).CellBackColor = RGB(153,204,255) 'bleu
	next i

'en-têtes colonnes planètes
	'ligne 1 : écriture en array des planètes
	for i=0 to 15
		zz(0,i)=planete(i)
	next i
	'ligne 1 : écriture en-têtes
	cell=Sheet.getCellrangeByPosition(1,0,16,0)
	cell.setdataarray(zz)
 
'en-têtes lignes (signes)
	'écriture en array des signes
	redim yy(11,0)
	for j=0 to 11
		yy(j,0)=signe(j)
	next j
	redim preserve yy(0 to 11,0)
	'écriture en-têtes à gauche colonne 0
	cell=Sheet.getCellrangeByPosition(0,1,0,12)
	cell.setdataarray(yy)
	'écriture en-têtes à droite colonne 33
	cell=Sheet.getCellrangeByPosition(17,1,17,12)
	cell.setdataarray(yy)

'message
	Cell = Sheet.getCellrangeByPosition(1,13,4,13)
	cell.Merge( True )
	Sheet.getCellByPosition(1,13).string="entrer les valeurs sous la forme 10.20"
		
'***partie médiane (aspects)***
	
'couleur des en-têtes H et V
	Sheet = Doc.Sheets.getByName("thème")
	sheet.getCellRangeByPosition(0,15,17,15).CellBackColor = RGB(153,204,255) 'Horizontal bleu 
	sheet.getCellRangeByPosition(0,15,0,31).CellBackColor = RGB(153,204,255) 'Vertical gauche bleu
	sheet.getCellRangeByPosition(17,15,17,31).CellBackColor = RGB(153,204,255) 'Vertical droite bleu
	sheet.getCellRangeByPosition(0,32,17,52).CellBackColor = RGB(224,224,224) 'zone inférieure (dignités, axes) gris
	sheet.getCellRangeByPosition(0,13,17,14).CellBackColor = RGB(224,224,224) ' ligne intermédiaire gris

'explication de 1ère colonne et 1ère ligne
	Sheet.getCellByPosition(0,14).string="sujet"
	Sheet.getCellByPosition(1,14).string="agent"

'en-têtes colonnes planètes
	'ligne 14 : écriture en array des planètes
	for i=0 to 15
		zz(0,i)=planete(i)
	next i
	redim preserve zz(0,15)
	cell=Sheet.getCellrangeByPosition(1,15,16,15)
	cell.setdataarray(zz) 
	
'en-têtes lignes planetes (gauche et droite)
	'écriture en array des planètes
	redim yy(15,0)
	for j=0 to 15
		yy(j,0)=planete(j)
	next j
	'écriture en-têtes à gauche colonne 0
	cell=Sheet.getCellrangeByPosition(0,16,0,31)
	cell.setdataarray(yy)
	'écriture en-têtes à droite colonne 33
	cell=Sheet.getCellrangeByPosition(17,16,17,31)
	cell.setdataarray(yy)
	
'en-têtes lignes supplémentaires de Saturne à Pluton
	for i = 33 to 36
		Cell = Sheet.getCellByPosition(0, i)
		cell.string= planete(i-27)
	next i	
	Sheet.getCellrangeByPosition(0,33,0,36).CellBackColor = RGB(102, 255, 178) 'vert clair

'en-têtes pour actants
	Sheet.getCellByPosition(7, 48).string="actant"
	Sheet.getCellByPosition(8, 48).string="score"
	Sheet.getCellByPosition(9, 48).string="masque"
	Sheet.getCellByPosition(10, 48).string="aspect"
	
	Sheet.getCellByPosition(7, 49).string="Soleil"
	Sheet.getCellByPosition(7, 50).string="Lune"
	Sheet.getCellByPosition(7, 51).string="Saturne"	
	
'lignes noires autour des cellules
	aBorder.Color = RGB(0,0,0) 'RGB(0,204,204) bleu clair
	aBorder.InnerLineWidth = 0
	aBorder.OuterLineWidth = 10
	aBorder.LineDistance = 0
	oBorder.LeftLine = aBorder
	oBorder.TopLine = aBorder
	oBorder.RightLine =aBorder
	oBorder.BottomLine = aBorder
	oBorder.isHorizontalLineValid =true
	oBorder.HorizontalLine =aBorder
	oBorder.isVerticalLineValid =true
	oBorder.VerticalLine =aBorder
	'tracé des lignes
	sheet.getCellRangeByPosition(0,0,17,12).TableBorder = oBorder 'tableau thème
	sheet.getCellRangeByPosition(0,15,17,31).TableBorder = oBorder 'tableau aspects
	sheet.getCellRangeByPosition(0,33,2,37).TableBorder = oBorder 'aspects au transpersonnelles
	sheet.getCellRangeByPosition(0,39,1,50).TableBorder = oBorder 'maisons en signes
	sheet.getCellRangeByPosition(3,38,6,51).TableBorder = oBorder  'maisons habitées et gouvernées
	sheet.getCellRangeByPosition(7,48,10,51).TableBorder = oBorder 'actants
	
'zones vertes en bas
	sheet.getCellRangeByPosition(0,33,0,37).CellBackColor = RGB(102, 255, 178) 'aspects au transpersonnelles
	sheet.getCellRangeByPosition(0,39,1,50).CellBackColor = RGB(102, 255, 178) 'maisons en signes
	sheet.getCellRangeByPosition(3,39,6,51).CellBackColor = RGB(102, 255, 178) 'maisons habitées et gouvernées
	sheet.getCellRangeByPosition(7,49,10,51).CellBackColor = RGB(102, 255, 178) 'actants
	
'""""alignement du texte au centre'''''
	sheet.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	'ajustement de la largeur des colonnes
	sheet.charheight=6
	for i=0 to 33
	Sheet.columns(i).Optimalwidth = True
	next i

End sub

Sub feuille_theme_progresse
Dim aBorder as New com.sun.star.table.BorderLine
Dim oBorder as New com.sun.star.table.TableBorder
dim nom as string

Doc = ThisComponent

'nom du thème
Sheet = Doc.Sheets.getByName("éphémérides")
nom=Sheet.getCellByPosition(19, 0).string

'effacement feuille progressé
If Doc.Sheets.hasByName("psychogenèse") Then Doc.Sheets.RemoveByName("psychogenèse")
	
'création feuille psychogenèse
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
 	Doc.Sheets.insertByName("psychogenèse", Sheet)
 	
Sheet = Doc.Sheets.getByName("psychogenèse")
'***partie inférieure (thème progressé)***
'en-tête général
	Cell = Sheet.getCellrangeByPosition(0,0,14,0) '(0,51,14,51)
	cell.Merge( True )
	Sheet.getCellByPosition(0,0).string=nom & " - tableau progressé + phases de lunaison progressée (lp) + phases génériques de Jupiter à Pluton (" & chr$(966) & ")"
	Sheet.getCellByPosition(0,0).HoriJustify = com.sun.star.table.CellHoriJustify.CENTER

'couleur des en-têtes H et V ligne 52 + à partir de la ligne 64
	sheet.getCellRangeByPosition(0,1,14,1).CellBackColor = RGB(153,204,255) 'Horizontal bleu 
	for i=13 to 112 step 11 'en-tête + lignes intermédiaires
	sheet.getCellRangeByPosition(0,i,14,i).CellBackColor = RGB(153,204,255) 'Horizontal bleu 
	next i
	sheet.getCellRangeByPosition(0,1,0,112).CellBackColor = RGB(153,204,255) 'Vertical gauche bleu
	sheet.getCellRangeByPosition(14,1,14,112).CellBackColor = RGB(153,204,255) 'Vertical droite bleu
	sheet.getCellRangeByPosition(0,0,14,0).CellBackColor = RGB(224,224,224) ' ligne intermédiaire gris
	
'en-tête colonne age
	Cell = Sheet.getCellByPosition(0,1) : cell.string= "âge"
'en-tête planètes (ligne 52)
	for i=1 to 12
	Sheet.getCellByPosition(i,1).string= planete(i-1)
	next i
	Sheet.getCellByPosition(13,1).string="NS"
	Sheet.getCellByPosition(14,1).string="AS"
	
'en-têtes colonnes planètes, répétés toutes les 10 lignes (à partir de la ligne 64)
for j=0 to 9 'nombre de lignes
	for i= 1 to 12 'colonnes
	Sheet.getCellByPosition(i, 13 + (11*j)).string= planete(i-1)
	next i
	Sheet.getCellByPosition(0,13 + (11*j)).string= nom
	Sheet.getCellByPosition(13,13 + (11*j)).string= "NS"
	Sheet.getCellByPosition(14,13 + (11*j)).string= "AS"
next j
'en-têtes lignes années avec saut d'une ligne toutes les dix lignes
  	car=0
	for i= 3 to 102
	Cell = Sheet.getCellByPosition(0, i+int(car/10)) : cell.value=i-3
	Cell = Sheet.getCellByPosition(14, i+int(car/10)) : cell.value=i-3
	car=car+1
	next i
	
'lignes noires autour des cellules
	aBorder.Color = RGB(0,0,0) 'RGB(0,204,204) bleu clair
	aBorder.InnerLineWidth = 0
	aBorder.OuterLineWidth = 10
	aBorder.LineDistance = 0
	oBorder.LeftLine = aBorder
	oBorder.TopLine = aBorder
	oBorder.RightLine =aBorder
	oBorder.BottomLine = aBorder
	oBorder.isHorizontalLineValid =true
	oBorder.HorizontalLine =aBorder
	oBorder.isVerticalLineValid =true
	oBorder.VerticalLine =aBorder
	'tracé des lignes
	sheet.getCellRangeByPosition(0,1,14,112).TableBorder = oBorder 'tableau thème progressé

'""""alignement du texte au centre'''''
'	sheet.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	'ajustement de la largeur des colonnes
	sheet.charheight=6
	for i=0 to 33
	Sheet.columns(i).Optimalwidth = True
	next i
	
End sub


Sub calcul_des_aspects

Dim aSortFields(0) As New com.sun.star.util.SortField 'pour trier colonnes
Dim aSortDesc(0) As New com.sun.star.beans.PropertyValue 'pour trier colonnes
dim actant%
dim annee_naissance as integer
dim colonne%
dim inc_col%, inc_lig%
dim ligne% 
dim coul as long
dim coul1%, coul2%
dim datemin as long
dim datemax as long
dim dignite1, dignite2 as string
dim ecart_heure  as double
dim finligne as string
dim gap as double
dim heure_naissance as double
dim index_planete as integer
dim maison as string
dim ns as double
dim offset as double
dim orbemax as double 'currency 'integer
dim phase as integer
dim phase_progressee as string
dim planete_R as string
dim quarante_ans_ok%
dim rangmin as integer
dim rangmax as integer
dim total as integer
dim val1 as integer

dim aa,bb,cc,dd,ee, eee 'arrays
dim ecart(1 to 12) as double
dim exaltation(2) as integer
dim exil(2) as integer
dim indice(2) as integer
dim maisons(0 to 11) as string
dim natal(0 to 15) as double 'longitudes thème natal
dim orbe_as_fc 'array pour orbes AS à MC


Doc = ThisComponent

If not Doc.Sheets.hasByName("éphémérides") Then msgbox "pas de feuille éphémérides, exécuter ephemerides" : exit sub

'confirmation calcul des aspects
	choix=inputbox ("caclcul des aspects du thème ?  (O,N)", "aspects","O")
	if choix <> "O" then exit sub

'désactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(false)

'inutile, appelé dans feuille_theme
'call definitions

'suppression et recréation feuille thème
	If Doc.Sheets.hasByName("thème") Then Doc.Sheets.RemoveByName("thème")
	call feuille_theme

Sheet = Doc.Sheets.getByName("éphémérides")
'récupération des longitudes du thème natal
	for i=0 to 15
	natal(i)=Sheet.getCellByPosition(i+1, 1).getvalue
	next i

'récupération des orbes page éphémérides pour AS à MC
	cell=Sheet.getCellRangeByPosition(17,1,17,16)
	orbe_as_fc=cell.getdata


'début

'recherche des aspects du thème
Sheet = Doc.Sheets.getByName("thème")	

	'planète sujet à comparer
	for i =0 to 15
			'remise à 0 de ecart(1 à 12) - sert à déterminer le masque des actants
			for j=1 to 12
			ecart(j)=0
			next j
		'planète agent
		for j=0 to 15
			'on ne compare pas une planète à elle-même	
			if j=i then goto fin_boucle
			'différence de 2 longitudes
			gap=natal(j)-natal(i)
			if gap < 0 then gap=gap+360
				'écriture phase
				cell =Sheet.getCellByPosition(j+1, i+16)
				phase=int(gap/30)+1	
				cell.string=chr(966) & str(phase)
			'division par 15 pour une approximation de l'aspect
			val1=int(gap/15)
			
						
			'aspects proches de cette approximation			
			for k=arc(val1,0) to arc(val1,1)
				'orbe à comparer à l'orbe maximum de la transitée
				orbedecimal=gap-angle(k)
			
				'utilisation des orbes définis dans orbe_theme(,,) de Soleil à Pluton et des orbes de la feuille éphémérides de AS à MC
				'déinition de l'orbe
				if i < 12 and j < 12 then
					orbemax=orbe_theme(i,j,k mod 16)
				else
					orbemax = orbe_as_fc(k mod 16)(0)
				endif
			
			 '******aspect trouvé*******
			 if abs(orbedecimal) <= orbemax then 
			 	abc=aspect(k mod 16)
			
				'récupération des phases des aspects pour les actants Soleil, Lune et Saturne (i=0,1,6) par rapport à agents Soleil à Pluton (j<10)
				if j <10 then if i=0 or i=1 or i=6 then
					'phase de l'aspect
					car=val(mid$(abc,instr(1,abc,chr(966))+1))
					'opposition=conjonction
					if car=7 then car=1
					'phases décroissantes=phases crosissantes
					if car >7 then car=14-car
					'multiplication par 100 pour bien séparer les phases
					car=car*100
					'ajout de l'orbe et sauvegarde
					ecart(j+1)=car+abs(orbedecimal)
				endif
				
				
				'écriture aspect
				cell =Sheet.getCellByPosition(j+1, i+16)
					'mise à la ligne de la phase de l'aspect
					car3=instr(1,abc," ")
					if car3 then
						cell.string=mid$(abc,1,car3-1) & chr$(10) & mid$(abc,car3+1)
						'car particulier de la conjonction phase 12
						if orbedecimal <0 and instr(1,abc,"conjonction") then cell.string="conjonction" & chr$(10) & chr(966) & "12"
					else
						cell.string=abc 
					endif
					
						
				'couleur
				cell.charcolor=couleur(k mod 16)
					'teste si conjonction arrière, dans ce cas mis en rouge sauf Mercure ou Vénus en conjonction au Soleil
					if k mod 16 =0 and orbedecimal <= 0 then
						coul=RGB(255,0,0) 'rouge
						bcd = planete(i) 'planète sujet
						cde = planete(j) 'planète agebt
							select case bcd
							case "Soleil"
							if cde ="Mercure" or cde ="Vénus" then coul=RGB(0,0,255) 'bleu
							case "Mercure"
							if cde ="Soleil" or cde ="Vénus" then coul=RGB(0,0,255)'bleu
							case "Vénus"
							if cde ="Soleil" or cde ="Mercure" then coul=RGB(0,0,255)'bleu
							end select
						cell.charcolor= coul 
					endif
			endif
		next k
fin_boucle:
	next j
	
	'recherche du masque de l'actant (si planète=Soleil, Lune ou Saturne)
	if i=0 or i=1 or i=6 then
		'valeur de départ 
		car=100000
		'recherche de l'aspect le plus court
		for k=1 to 9
			if ecart(k) >0 and ecart(k) < car then car=ecart(k) : car1=k
		next k
		'ligne où écrire le masque
		car2=49+i
			'Saturne
			if car2=55 then car2=51
		'écriture du masque
		Sheet.getCellByPosition(9, car2).string=planete(car1-1)
		'écriture aspect du masque
		sheet.getcellbyposition(10,car2).string=sheet.getcellbyposition(car1,16+i).string
	endif
			
next i



'calcul du nombre d'aspects aux trans-personnelles
Sheet = Doc.Sheets.getByName("thème")
	'de Saturne à Pluton en horizontal
	for j=7 to 10
	car=0 : coul1=0 : coul2=0
		'du Soleil à Jupiter en vertical
		for i= 16 to 21
		cell=Sheet.getCellByPosition(j, i)
			'si aspect présent (couleur n'est pas noire -1)
			if cell.charcolor <>-1 then
			car=car+1
			total=total+1
				select case cell.charcolor
				case RGB(0,0,255),RGB(153, 51, 255) 'bleu ou mauve (quinconce)
				coul1=coul1+1
				case RGB(255,0,0),RGB(255,51,255) 'rouge ou violet (sesqui-carré)
				coul2=coul2+1
				end select
			endif
		next i
		
		'nombre des aspects d'une planète (+ détail)
		'planète + total aspects
		Cell = Sheet.getCellByPosition(0,26+j) 
		cell.string=cell.string & " " & str(car)
		cell.HoriJustify = com.sun.star.table.CellHoriJustify.RIGHT
		'total bleus par planète
		Cell = Sheet.getCellByPosition(1,26+j) 
		cell.charcolor=RGB(0,0,255)
		cell.value=coul1
	'	cell.HoriJustify = com.sun.star.table.CellHoriJustify.RIGHT
		'total bleus
		Cell = Sheet.getCellByPosition(1,37) 
		cell.value=cell.value+coul1
		cell.charcolor=RGB(0,0,255)
	'	cell.HoriJustify = com.sun.star.table.CellHoriJustify.RIGHT
		'total rouges par planète
		Cell = Sheet.getCellByPosition(2,26+j) 
		cell.charcolor=RGB(255,0,0)
		cell.value=coul2
	'	cell.HoriJustify = com.sun.star.table.CellHoriJustify.LEFT
		'total rouges
		Cell = Sheet.getCellByPosition(2,37) 
		cell.value=cell.value+coul2
		cell.charcolor=RGB(255,0,0)
	'	cell.HoriJustify = com.sun.star.table.CellHoriJustify.LEFT
	next j


'total des aspects aux transpersonnelles
	cell= Sheet.getCellByPosition(0, 37)
	cell.string="total" & " " & str(total)
	cell.HoriJustify = com.sun.star.table.CellHoriJustify.RIGHT

'récupération du nom de fichier ex "astro.txt"
	Sheet = Doc.Sheets.getByName("éphémérides")
	abc = Sheet.getCellByPosition(19, 0).getstring
	Sheet = Doc.Sheets.getByName("thème")
	'partie thème
	Cell = Sheet.getCellByPosition(0, 0) : cell.string=abc
	Cell.CellBackColor = RGB(0, 255, 0) 'vert clair
	Cell = Sheet.getCellByPosition(33, 0) : cell.string=abc
	Cell.CellBackColor = RGB(0, 255, 0) 'vert clair
	'partie aspects
	Cell = Sheet.getCellByPosition(0, 15) : cell.string=abc
	Cell.CellBackColor = RGB(0, 255, 0) 'vert clair
	Cell = Sheet.getCellByPosition(17, 15) : cell.string=abc
	Cell.CellBackColor = RGB(0, 255, 0) 'vert clair

'écriture date et heure de naissance
	'lecture feeuille éphémérides
	Sheet = Doc.Sheets.getByName("éphémérides")
		'date de naissance au format numérique
		date_naissance=Sheet.getCellByPosition(22, 1).getvalue
		'âge > 40 ans ?
		car=40*365
		if datevalue(now)-date_naissance > car then quarante_ans_ok=1
		'heure de naissance au format string
		abc=Sheet.getCellByPosition(22, 2).getstring
	'écriture feuille thème
	Sheet = Doc.Sheets.getByName("thème")
		'date de naissance
		cell=sheet.getCellByPosition(5, 13)
		cell.value=date_naissance
	   'format dd/mm/aaaa
		cell.numberformat=10030
		'heure de naissance
		sheet.getCellByPosition(6, 13).string=abc
		
'ajustement de la largeur des colonnes
	Sheet = Doc.Sheets.getByName("thème")
	for i=0 to 33
	Sheet.columns(i).Optimalwidth = True
	next i


''''''''remplissage du thème à partir des données page éphémérides + calcul des dignités et des axes'''''''''


'dignités
'en-têtes des dignités
	Sheet.getCellByPosition(3,33).string = "Maîtrise"
	Sheet.getCellByPosition(3,34).string = "Exil"
	Sheet.getCellByPosition(3,35).string = "Chute"
	Sheet.getCellByPosition(3,36).string = "Exaltation"
	Sheet.getCellByPosition(3,37).string = "Rétrograde(s)"
'couleur vert clair + alignement
	cell=sheet.getCellRangeByPosition(3,33,3,37)
	cell.CellBackColor = RGB(102, 255, 178) 'vert clair

'axes
'en-têtes croix des axes + couleurs
	cell=Sheet.getCellByPosition(11,40) : cell.string = "MC" : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
	cell= Sheet.getCellByPosition(10,41) : cell.string = "AS" : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
	cell=Sheet.getCellByPosition(12,41) : cell.string = "DS" : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
	cell=Sheet.getCellByPosition(11,42) : cell.string = "FC" : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
	Sheet.getCellByPosition(11,39).CellBackColor = RGB(102, 255, 178) 'vert clair
	Sheet.getCellByPosition(11,43).CellBackColor = RGB(102, 255, 178) 'vert clair
	Sheet.getCellByPosition(9,41).CellBackColor = RGB(102, 255, 178) 'vert clair
	Sheet.getCellByPosition(13,41).CellBackColor = RGB(102, 255, 178) 'vert clair


'lecture et conversion des données du thème de longitudes en degrés minutes
for index_planete=1 to 16
	'lecture longitudes feuille éphémérides
	Sheet = Doc.Sheets.getByName("éphémérides")
	planete_R=Sheet.getCellByPosition(index_planete, 0).getstring 'nom planète avec (R) éventuellement
	cell=Sheet.getCellByPosition(index_planete, 1)
	'conversion en degrés/minutes
 	call calc_deg_min(cell,0)
	
'écriture feuille thème
	Sheet = Doc.Sheets.getByName("thème")
	'degrés, minutes
	cell=Sheet.getCellByPosition(index_planete,index_signe + 1)
	cell.string=str(degres) & chr$(176) & str(minutes) & "'"
	Cell.NumberFormat = 10108
	'dignités de Soleil à Neptune
	if index_planete < 11 then
		dignite1="" : dignite2=""
		abc=matrice(index_planete-1,index_signe)'dignité
		car=instr(1,abc, "et")
		'double dignité pour Mercure
		if car then
			dignite1=mid$(abc,1,car-2)
			dignite2=mid$(abc,car+3)
		else
			dignite1=abc
		endif
		'écriture et comptage des dignités
		abc=dignite1
		for i=1 to 2
			
			select case abc
			case "Maîtrise"
			cell=Sheet.getCellByPosition(4,33) : cell.value=cell.value+1 : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
				car=5
				do until Sheet.getCellByPosition(car,33).getstring=""
				car=car+1
				loop
				cell=Sheet.getCellByPosition(car,33) : cell.string=planete(index_planete-1) : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
				
			case "Exil"
			cell=Sheet.getCellByPosition(4,34) : cell.value=Sheet.getCellByPosition(4,34).value+1 : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
				car=5
				do until Sheet.getCellByPosition(car,34).getstring=""
				car=car+1
				loop
				cell=Sheet.getCellByPosition(car,34) : cell.string=planete(index_planete-1) : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
				
			case "Chute","Chute-","Chute+"
			cell=Sheet.getCellByPosition(4,35) : cell.value=Sheet.getCellByPosition(4,35).value+1 : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
				car=5
				do until Sheet.getCellByPosition(car,35).getstring=""
				car=car+1
				loop
				cell=Sheet.getCellByPosition(car,35) : cell.string=planete(index_planete-1) : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
				'chute vertueuse (chute+), dépravee (chute-) ou neutre(chute) ?
				if instr(1,abc,"+") then cell.string= cell.string & "+"
				if instr(1,abc,"-") then cell.string= cell.string & "-"
				
			case "Exaltation"
			cell=Sheet.getCellByPosition(4,36) : cell.value=Sheet.getCellByPosition(4,36).value+1 : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
				car=5
				do until Sheet.getCellByPosition(car,36).getstring=""
				car=car+1
				loop
				cell=Sheet.getCellByPosition(car,36) : cell.string=planete(index_planete-1) : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
			end select
			
			abc=dignite2
		next i
	endif
		
		'rétrogradation
		if instr (1,planete_R,"(R)") then
		cell=Sheet.getCellByPosition(4,37) : cell.value=Sheet.getCellByPosition(4,37).value+1 : cell.CellBackColor = RGB(102, 255, 178) 'vert clair
				car=5
				do until Sheet.getCellByPosition(car,37).getstring=""
				car=car+1
				loop
				cell=Sheet.getCellByPosition(car,37) : cell.string=planete(index_planete-1) : cell.CellBackColor = RGB(102, 255, 178) 'vert clair	
		endif
		
		'signes angulaires de AS à MC
		if index_planete > 12 and index_planete < 17 then
			'abc : nom planète, bcd : signe
			abc=Sheet.getCellByPosition(index_planete,0).getstring
			bcd=Sheet.getCellByPosition(0,index_signe+1).getstring 
			select case abc
			case "AS"
			Sheet.getCellByPosition(9,41).string=bcd
			case "FC"
			Sheet.getCellByPosition(11,43).string=bcd
			case "DS"
			Sheet.getCellByPosition(13,41).string=bcd
			case "MC"
			Sheet.getCellByPosition(11,39).string=bcd 
			end select
		endif
next index_planete

'aspect entre AS et FC
	cell=Sheet.getCellByPosition(11,41) : cell.string=Sheet.getCellByPosition(13, 29).getstring 
	cell.HoriJustify = com.sun.star.table.CellHoriJustify.LEFT

'planètes conjointes aux axes
	'AS à MC en horizontal
	for i=13 to 16
		'Soleil à NN 'Pluton en vertical
		for j=16 to 27 '25
		cell=Sheet.getCellByPosition(i,j)
		abc= cell.string
			'if abc = "conjonction" then
			if instr(1,abc,"conjonction") then
				coul=cell.charcolor
				select case Sheet.getCellByPosition(i,15).string
				case "AS"
				ligne=42 : inc_lig=1 : colonne=9 : inc_col=0
				maison=" I"
				if coul=RGB(0,0,255) then maison = " XII" 'bleu
				case "FC"
				ligne=44 : inc_lig=1 : colonne=11 : inc_col=0
				maison=" IV"
				if coul=RGB(0,0,255) then maison = " III" 'bleu
				case "DS"
				ligne=42 : inc_lig=1 : colonne=13 : inc_col=0
				maison=" VII"
				if coul=RGB(0,0,255) then maison = " VI" 'bleu
				case "MC"
				ligne=39 : inc_lig=0 : colonne=12 : inc_col=1
				maison=" X"
				if coul=RGB(0,0,255) then maison = " IX" 'bleu
				end select
					'recherche cellule vide
					do until Sheet.getCellByPosition(colonne,ligne).getstring=""
					colonne=colonne+inc_col
					ligne=ligne+inc_lig
					loop
				'écriture planète + couleur
				cell=Sheet.getCellByPosition(colonne,ligne)
				cell.string=planete(j-16) & maison
				cell.charcolor= coul 'bleu ou rouge
			endif
		next j
	next i

'maisons en signes
for i=1 to 12
	Sheet.getCellByPosition(0,i+38).string="Maison" & str(i)
	'récupération position maison dans feuille éphémérides
Sheet = Doc.Sheets.getByName("éphémérides")
	car= Sheet.getCellByPosition(21, i).getvalue
	'écriture signe correspondant dans feuille thème
Sheet = Doc.Sheets.getByName("thème")
	if car then	car1=int(car/30) : Sheet.getCellByPosition(1,i+38).string=signe(car1)
next

'recherche des signes orphelins (interceptions)
	'signes
	for j=0 to 11
	'car=1 si signe trouvé
	car=0
		for i = 39 to 50 'signes réels des maisons
		abc=Sheet.getCellByPosition(1,i).getstring
		if abc=signe(j) then car=1 : exit for
		next
	'signe orphelin, mémorisé
		if car=0 then maisons(j) = signe(j)
	next

'ajout des signes orphelins avec un /
for j = 0 to 11 'liste des signes orphelins (sans maisons)
	if maisons(j) <> "" then
		car=j
		if j=0 then car=12 'si bélier on recherche poissons (correction passé car de 11 à 12 sinon Bélier/Verseau aulieu de Bélier/Poissons !)
		'recherche du signe précédant l'orphelin
		for i=50 to 39 step -1
		abc=Sheet.getCellByPosition(1,i).getstring
			'trouvé le signe précédant l'orphelin
			if abc=signe(car-1) then 
			car1=i
				'si maisons 1 et 12 de même signe, le signe orphelin va en maison 1 (ligne 39), pas en maison 12 (ligne 50)
				if i=50 then if Sheet.getCellByPosition(1,39).getstring=abc then car1=39
			'écriture
			Sheet.getCellByPosition(1,car1).string=abc & "/" & maisons(j)
			exit for
			endif
		next
	endif	
next


'ajout NS ligne 52 (+ signe)
	'lecture NN page éphémérides et ajout de 180
Sheet = Doc.Sheets.getByName("éphémérides")
	ns=Sheet.getCellByPosition(11, 1).getvalue + 180
	if ns >=360 then ns=ns-360
	'calcul du signe
	coeff1=int(ns/30)*30
	abc=signe(coeff1/30)
  'écriture page thème
Sheet = Doc.Sheets.getByName("thème")
	Sheet.getCellByPosition(3,51).string="NS"
	'signe NS
	Sheet.getCellByPosition(4,51).string=abc


'planètes en maisons habitées
	Sheet.getCellByPosition(5,38).string="habitée"
	'Lune à NS 'Lilith
	for i= 0 to 12 '11
		maison=""
		'écriture nom planète sauf NS déjà écrit (i=12)
		if i < 12 then Sheet.getCellByPosition(3,i+39).string=Sheet.getCellByPosition(0,i+16).getstring
		'récupération position planète dans feuille éphémérides
Sheet = Doc.Sheets.getByName("éphémérides")
		'récupération maison habitée à partir de la feuille éphémérides
			'maisons
			for j=1 to 12 
			car= Sheet.getCellByPosition(i+1, 1).getvalue
				'longitude NS
				if i=12 then car=ns
			'cuspide en cours
			car1=Sheet.getCellByPosition(21, j).getvalue
			if car1=0 then maison="?": exit for
			'cuspide suivante
			car2=Sheet.getCellByPosition(21, j+1).getvalue
				'cas de rebouclage (passage de 360 à 0)
				if car1-car2 > 300 then
				car2=car2+360
					if car1-car > 300 then car=car+360
				endif
				'recherche
				if car > car1 and car < car2 then
				maison=Sheet.getCellByPosition(20, j).getstring
				exit for
				endif
			next
			'si pas trouvé, Maison XII obligatoirement
			if maison="" then maison="12"
			
		'écriture signe et maison habitée dans feuille thème
Sheet = Doc.Sheets.getByName("thème")
		if car > 360 then car=car-360 'suppression de l'ajout de 360 en cas de rebouclage (v. plus haut)
		car1=int(car/30)
		Sheet.getCellByPosition(4,i+39).string=signe(car1)
		Sheet.getCellByPosition(5,i+39).string=maison
	next
	

	
'planètes et maisons gouvernées
Sheet.getCellByPosition(6,38).string="gouvernée(s)"
'Lune à Pluton
for i =0 to 9
	for j=0 to 11 'recherche des maîtrises dans matrice() v. deinitions
		if instr (1,matrice(i,j),"Maîtrise") then
		abc=signe(j) 'signe de maîtise pour cette planète
			'recherche du signe correspondant à la maîtrise
			for k= 39 to 50 'lignes des maisons feuille thème
			'recherche maisons gouvernées
			 cde=Sheet.getCellByPosition(1,k).getstring 'signe maison
				 if instr(1,cde,abc) then
					bcd=Sheet.getCellByPosition(0,k).getstring 'maison 1, maison 2, etc.
					cde=Sheet.getCellByPosition(6,i+39).getstring 'numéro maison gouvernée si existe déjà
					finligne=""
					if cde <> "" then finligne=", "
					'écriture
					Sheet.getCellByPosition(6,i+39).string=cde & finligne & RIGHT(bcd, 2) 'numéro maison gouvernée
				endif 
			next
		endif
	next
next

'calcul phase Soleil-Lune
Sheet = Doc.Sheets.getByName("éphémérides")
		phase_progressee=""
		gap=Sheet.getCellByPosition(2,1).getvalue-Sheet.getCellByPosition(1,1).getvalue 'lune-soleil
		if gap<0 then gap=gap+360
		gap=int(gap/30)+1
		phase_progressee=chr$(966) & str(gap)

Sheet = Doc.Sheets.getByName("thème")
		Sheet.getCellByPosition(3,38).string="Soleil-Lune " & phase_progressee


'ordre des transits de la Lune progressée sur les planètes
Sheet = Doc.Sheets.getByName("éphémérides")
	'effacement colonnes X,Y,Z,AA (sinon erreur à la 2ème lecture)
	for i=1 to 13
		for j=23 to 26
		sheet.getcellbyposition(j,i).string=""
		next j
	next i
	'copie des positions planètes colonnes X (nom),Y (longitude), et Z (index) (23,24,25)
		for i=1 to 12
		'noms planètes
			'retire si besoin (R) du nom
			abc=Sheet.getCellByposition(i,0).string
			car=instr(1,abc,"(")
			if car then abc=mid$(abc,1,car-2)
		Sheet.getCellByposition(23,i).string=abc
		'longitudes
		Sheet.getCellByposition(24,i).value=Sheet.getCellByposition(i,1).value
		'index planètes
		Sheet.getCellByposition(25,i).value=i-1
		next i
	'ajoute Noeud Sud NS=NN+180
		'nom
		Sheet.getCellByposition(23,13).string="NS"
		'longitude
		Sheet.getCellByposition(24,13).value=Sheet.getCellByposition(24,11).value+180
		if Sheet.getCellByposition(24,13).value>=360 then Sheet.getCellByposition(24,13).value=Sheet.getCellByposition(24,13).value-360
		'index
		Sheet.getCellByposition(25,13).value=12
	'tri par ordre croissant des 3 colonnes X,Y,Z
	cell=sheet.getcellrangebyposition(23,1,25,13)
	aSortFields(0).Field = 0
	aSortFields(0).SortAscending = true
	'(field = 0 : tri sur 1ère colonne, field = 1 : tri sur 2ème colonne)
	aSortFields(0).Field = 1 
	aSortFields(0).SortAscending = true
	aSortDesc(0).Name = "SortFields"
	aSortDesc(0).Value = aSortFields()
	cell.Sort(aSortDesc())
	'car = ligne Lune
	for i=1 to 13
	if Sheet.getCellByposition(23,i).string="Lune" then car=i : exit for
	next i
	'ajout colonne AA (26) du numéro d'ordre de 0 à 12 (car1) en face de chaque planète (ordre de transit de la lune sur la planete) (Lune=0 1ère ligne car)
	car1=0
	for i=1 to 13
	Sheet.getCellByposition(26,car).value=car1
	car1=car1+1
	car=car+1
	if car=14 then car=1
	next i
	
	'lecture données feuille éphémérides colonne X (planète), Z (ordre) et AA (index)
	for i=1 to 13
Sheet = Doc.Sheets.getByName("éphémérides")
		'planète
		abc= sheet.getcellbyposition(23,i).string
		'car1 : numéro d'ordre du transit lune progressée
		car1=sheet.getcellbyposition(26,i).value
		'car2: numéro d'index de la planète
		car2=sheet.getcellbyposition(25,i).value
		'écriture feuille thème lignes 39 à 52 colonne D
Sheet = Doc.Sheets.getByName("thème")
		cell=sheet.getcellbyposition(2,39+car2)
		cell.string=str(car1)
	next i
	'en-tête
	sheet.getcellbyposition(2,38).string="lp"
	sheet.getcellrangebyposition(2,38,2,51).HoriJustify = com.sun.star.table.CellHoriJustify.RIGHT



'actants : calcul du score (le masque est déterminé plus haut)
	Sheet = Doc.Sheets.getByName("thème")
	'remise à 0 des scores
	for i=0 to 2
		indice(i)=0
	next i
	'phases + aspects
		'lignes Soleil, Lune et Saturne
		aa=array(16,17,22)
		'indices des phases sans aspects
		bb=array(2,1,2,-2,2,-1,0,-3,2,-2,1,-2)
		'indices des phases avec aspects
		select case quarante_ans_ok
		'<40 ans
		case 0
			'indices phases
			cc=array(4,2,4,-4,4,-3,2,-3,4,-4,4,2,-3)
			'indices interphases
			dd=array(-3,2,2,-3)
		'>40 ans
		case 1
			'indices phases
			cc=array(4,1,4,-4,4,-3,-2,-3,4,-4,4,-3,-4)
			'indices interphases
			dd=array(-3,-2,-2,-3)
		end select
		
		'actants
		for i=0 to 2
		
			'coeficients planètes, à ajouter (ou multiplier ?) aux indices de phases
			select case i
			'Soleil
			case 0
				'phase sans aspect
				ee=array(0,0,0,0,1,2,3,3,3,3)
				'phase avec aspect
				eee=array(0,0,0,0,1,2,5,6,7,12) 
			'Lune
			case 1
				ee=array(3,0,1,2,4,4,4,4,4,4)
				eee=array(3,0,1,2,4,5,8,9,10,15)
			'Saturne
			case 2
				ee=array(0,0,0,0,0,0,0,1,2,3)
				eee=array(0,0,0,0,0,0,0,1,2,7)
			end select
			
			'agents Soleil à Pluton
			for j=1 to 10
				'récupération phases (sans ou avec aspect)
				cell=sheet.getcellbyposition(j,aa(i))
				abc=cell.string
					'cellule vide
					if abc="" then goto finji
				'phase
				car=instr(1,abc,chr(966))
				phase=val(mid$(abc,car+1))
					'phase sans aspect (couleur=noir)
					if cell.charcolor=-1 then
						indice(i)=indice(i)+(bb(phase-1)*ee(j-1))
					'phase avec aspect
					else
						'interphase ?
							'oui
							if mid$(abc,car-1,1)="i" then
							 	indice(i)=indice(i)+(dd(int(phase/3))*eee(j-1))
							 'non
							 else
							 	'conjonction phase 12 ?
							 	'oui
							 	if phase=12 and instr(1,abc,"conjonction") then
							 		indice(i)=indice(i)+(cc(12)*eee(j-1))
							 	'non
							 	else	
							 		indice(i)=indice(i)+(cc(phase-1)*eee(j-1))
							 	endif
							 endif	
					endif
				'ajout (ou multiplication ?) coefficient planète
				'indice(i)=indice(i)*ee(j-1)	
		finji:
			next j
		next i
		
	
		
	'dignités
		'maitrise, exil, chute, exaltation
		for i=33 to 36
			'car=nombre de planètes ayant cette dignité
			car=sheet.getcellbyposition(4,i).getvalue
			'aucune
			if car<=0 then goto finii
			
			for j=1 to car
				'planète en dignité
				abc=sheet.getcellbyposition(4+j,i).getstring
					'actant ?
					actant=-1
					if instr(1,abc,"Soleil") then actant=0
					if instr(1,abc,"Lune") then actant=1
					if instr(1,abc,"Saturne") then actant=2
				'actant : oui	
				if actant >=0 then
					select case i
					'maitrise
					case 33
						indice(actant)=indice(actant)+30
					'exil
					case 34	
						 indice(actant)=indice(actant)*-1
						 if indice(actant) >0 then exil(actant)=1 '=réintégration de l'actant
						 if indice(actant) >15 then indice(actant)=15
						 if indice(actant) <-15 then indice(actant)=-15
					'chute
					case 35
						'vertueux
						if instr(1,abc,"+") then
							indice(actant)=indice(actant)-30
						'dépravé
						elseif instr(1,abc,"-") then
							if actant=2 then indice(actant)=indice(actant)+15 else indice(actant)=indice(actant)-15
						'mutable
						else
							indice(actant)=indice(actant)-20
						endif
					'exaltation
					case 36
						exaltation(actant)=1
						if indice(actant) >0 then
							indice(actant)=indice(actant)*2
							if indice(actant) <30 then indice(actant)=30
						else
							indice(actant)=indice(actant)*-1
							if indice(actant) >30 then indice(actant)=30
						endif 
					end select
				endif
			next j
	finii:
		next i

		
	'maison habitée
		'lignes maisons habitées des actants
		bb=array(39,40,45)
		'coeff Soleil
		aa(0)=array(60,20,40,30,90,-30,0,-40,60,90,40,-60)
		'coeff Lune
		aa(1)=array(120,20,30,60,90,-60,0,-90,90,50,60,-120)
		'coeff Saturne
		aa(2)=array(30,10,20,-10,10,-20,0,-30,30,40,20,-40)
		
		'ajout du coeff au score
		for i=0 to 2
			'car=numéro maison habitée
			car=val(sheet.getcellbyposition(5,bb(i)).getstring)
			'maison habitée est bien définié
			if car then
				'Si exaltation et maison négative, Si (valeur actant+valeur maison) <0 Alors val actant/2+0
				if exaltation(i)=1 and aa(i)(car-1) <0 then if indice(i)=indice(i)+ aa(i)(car-1) <0 then indice(i)=indice(i)/2 : goto fin_maison
										
				'Si exil >0  (réintégration d’un actant) Si Valeur actant>0 et maison positive: Alors val actant+val maison/2
				if exil(i)=1 then if indice(i) >0 and aa(i)(car-1) >0 then indice(i)=indice(i)+ (aa(i)(car-1)/2) : goto fin_maison
				
				'Autres cas Alors val actant+val maison
				indice(i)=indice(i)+ aa(i)(car-1)
			endif
	fin_maison:
		next i


		
	'écriture du score feuille thème
		'lignes
		for i=49 to 51
			sheet.getcellbyposition(8,i).value=indice(i-49)
		next i	



		
'''''*****************************''choix de calcul du thème progressé et des phases génériquese'''''**********************************************''''

choix=inputbox ("calcul du thème progressé et des phases génériques ? (O,N)", "thème progressé et phases génériques","O")
if choix="" then goto fin
call feuille_theme_progresse
message=""

dim as_lp as string
dim long_lune as double
dim lpp as string, lpn as string
Dim position(0 to 99) as long 'pour phases génériques

'''''''tableau progressé de 1 à 90 jours=ans''''''''''
Sheet = Doc.Sheets.getByName("éphémérides")

'récupère le nombre de lignes utilisées dans la feuille éphémérides
	Sheet = Doc.Sheets.getByName("éphémérides")
	Curs = Sheet.createCursor
	Curs.gotoEndOfUsedArea(True)
	lastrow = Curs.Rows.Count

'vérifie si date de naissance présente page éphémérides colonne W
	abc=Sheet.getCellByPosition(22, 1).getstring
	'pas de date de naissance ou date à priori incorrecte

	if abc="" or len(abc) <> 10 then
	message= " - pas de date de naissance !, lancer calcul_du_theme"
	goto fin2
	endif
	
	annee_naissance=val(mid$(abc,7,4)) '+ 1 'démarrage à jour de naissance + 1 = année de naissance + 1
	'année de naissance incorrecte
	if annee_naissance < 1900 or annee_naissance > 2050 then
	message= " - année de naissance hors plage éphémérides, pas de calcul du tableau progressé" 
	goto fin2
	endif
	
'date en format numérique		
date_naissance=datevalue(abc)

'vérifie si date de naissance incluse dans les éphémérides
abc=Sheet.getCellByPosition(0, 18).getstring
	'pas d'éphémérides, sortie
	if abc="" then
	message=" - éphémérides absents, pas de calcul du tableau progressé" 
	goto fin2
	endif
	
datemin= datevalue(Sheet.getCellByPosition(0, 18).getstring)
datemax= datevalue(Sheet.getCellByPosition(0, lastrow-1).getstring)
	'date de naissance en-dehors des éphémérides
	if date_naissance < datemin or date_naissance > datemax then
	message= " - date de naissance hors plage éphémérides, pas de calcul du tableau progressé"
	goto fin2
	endif 
	
'recherche de la plage de données (1 à 90 jours après la date de naissance)
car=date_naissance '+1
for i= 17 to lastrow -1
	'trouvé date de naissannce + 1 jour
	datemin=datevalue(Sheet.getCellByPosition(0, i).getstring)
	if datemin=car then
	rangmin=i
	exit for
	endif
next i
if rangmin=0 then message= "erreur, pas de date trouvée" : goto fin2

rangmax=rangmin+99


'détermine l'écart de longitudes entre 2 jours successifs pour chaque planète (écart/12 = 1 mois en progressé, /360 = 1 jour)
 	'timevalue=coeff de 0 à 1 (midi=0,5)
	heure_naissance=timevalue(Sheet.getCellByPosition(22,2).getstring)
	for i=1 to 12
	ecart(i)=Sheet.getCellByPosition(i, rangmin+1).getvalue-Sheet.getCellByPosition(i, rangmin).getvalue
	'compensation si passage à 360 entre 2 lignes
	'planète rétrograde
	if (ecart(i)) > 300 then ecart(i)= ecart(i)-360
	'planète directe
	if (ecart(i)) < -300 then ecart(i)= ecart(i)+360
	'offset >0 ou <0 (planète (R)) pour tenir compte de l'heure de naissance
	ecart(i)=heure_naissance*(ecart(i))
	next i

'limites basses et hautes de la barre de progression (si form 'visuel' active)
	if form_ok=1 then
	commande1 = feuille.getControl("ProgressBar1")
	commande1.setrange(rangmin,rangmax)
	feuille.getControl("ProgressBar2").visible=false
	feuille.getControl("Label3").text="phases progressées"
	
	endif

'début
 	 
'lecture données de jour de naissance à jour+99 page éphémérides et écriture page thème
car3=0 'compteur pour saut de ligne toutes les 10 lignes

'lignes éphémérides
for i= rangmin to rangmax
	'barre de progression
	if form_ok=1 then commande1.setvalue(i)
	'Soleil à Lilith
	for j=1 to 12
		'lecture longitudes feuille éphémérides
Sheet = Doc.Sheets.getByName("éphémérides")
		cell=Sheet.getCellByPosition(j, i)
			'pour la ligne de l'année de naissance, lecture des données réelles, pas celles des éphémérides
			if i=rangmin then cell=Sheet.getCellByPosition(j, 1)
	'ajoute offset à la longitude sauf pour l'année de naissance
	'offset positif ou négatif (si planète rétrograde) pour l'heure de naissance (coefficient multiplicateur de 0 à 1 appliqué à l'écart de longitudes en 24h))
	offset=0
	if i <> rangmin then offset=ecart(j)
	'conversion en degrés/minutes
	call calc_deg_min(cell,offset)
		'si orange, planète rétrograde
			coul=Cell.CellBackColor 
		'calcul phases Lune progressée/planète progressée et natale (lpp et lpn)
			lpp="" : lpn=""
			long_lune=Sheet.getCellByPosition(2, i).getvalue
			'lpp (lunaison progressé/progressé)
			gap=long_lune + ecart(2) - longitude
			if gap<0 then gap=gap+360
			phase=int(gap/30)+1
			lpp="lpp" & str(phase)
			'lpn (lunaison progressé/natal)
			gap=long_lune + ecart(2) - natal(j-1)
			if gap<0 then gap=gap+360
			phase=int(gap/30)+1
			lpn="lpn" & str(phase)
		'calcul phase Lune progressée/AS natal
			gap=long_lune + ecart(2) - Sheet.getCellByPosition(13, 1).getvalue
			if gap<0 then gap=gap+360
			phase=int(gap/30)+1
			as_lp="lpn" & str(phase)
		'écriture feuille progressé
Sheet = Doc.Sheets.getByName("psychogenèse")
			'signe
				abc=signe(index_signe)
			'orange =(R)
				if coul= RGB(255, 213, 6) then abc=abc & " (R)"
			'degrés, minutes
				abc=abc & chr$(13) & str(degres) & chr$(176) & str(minutes) & "'"
			'phase progressée le cas échéant
				'pas de lpp pour la lune (j=2)
				if lpp <> "" and j <> 2  then abc=abc & chr$(13) & lpp
				if lpn <> "" then abc=abc & chr$(13) & lpn
			'écriture avec saut d'une ligne toutes les 10 lignes avec int(car3/10)
				cell=Sheet.getCellByPosition(j,3+i-rangmin+int(car3/10))
				cell.string=abc
				Cell.CellBackColor=coul
				'écriture phase AS/Lp
				cell=Sheet.getCellByPosition(14,3+i-rangmin+int(car3/10))
				cell.string=as_lp
			'ajoute année de naissance à l'âge dans la 1ère colonne
				if j=1 then
					abc=Sheet.getCellByPosition(0,3+i-rangmin+int(car3/10)).string
					abc=abc  & " - " & str(annee_naissance)
					Sheet.getCellByPosition(0,3+i-rangmin+int(car3/10)).string=abc
					annee_naissance=annee_naissance+1
				endif
	next j
car3=car3+1
next i

'recherche date de lunaison progressée prénatale
Sheet = Doc.Sheets.getByName("éphémérides")
cde=""
	for i=rangmin  to rangmin - 30 step -1
	'soleil-lune en valeur absolue
	position_epheref1=abs(Sheet.getCellByPosition(1,i).getvalue-Sheet.getCellByPosition(2,i).getvalue)
	position_epheref2=abs(Sheet.getCellByPosition(1,i-1).getvalue-Sheet.getCellByPosition(2,i-1).getvalue)
		'trouvé
		if position_epheref1 < 15 and position_epheref2 > position_epheref1 then
		cde=Sheet.getCellByPosition(0,i).getstring 'date de lunaison progressée prénatale
		ligne=i 'numéro de ligne
		exit for
		endif
	next i
	'lecture / écriture signes, deg/min
		'signes soleil et lune de la lunaison progressée prénatale
		for i=1 to 2
		'lecture
Sheet = Doc.Sheets.getByName("éphémérides")
		cell=Sheet.getCellByPosition(i, ligne)
		
	'conversion en degrés/minutes
	call calc_deg_min(cell,ecart(i))
		'signe + degrés/minutes 
		abc=signe(index_signe) & chr$(13) & str(degres) & chr$(176) & str(minutes) & "'"
		'écriture
		Sheet = Doc.Sheets.getByName("psychogenèse")
		Sheet.getCellByPosition(i,2).string=abc
		next i
	'lecture date de naissance au format string
Sheet = Doc.Sheets.getByName("éphémérides")
	abc=Sheet.getCellByPosition(0,rangmin).string
	'écriture	
Sheet = Doc.Sheets.getByName("psychogenèse")
	'nom
	cell=Sheet.getCellByPosition(0,1)
	cell.string=Sheet.getCellByPosition(0,13).string
	Cell.CellBackColor = RGB(0, 255, 0) 'vert clair
	'date de lunaison progressée prénatale
	Sheet.getCellByPosition(0,2).string="NL prénatale" & chr$(13) & cde 
	'date de naissance
	Sheet.getCellByPosition(0,3).string=abc
	'heure de naissance
	Sheet.getCellByPosition(0,3).string=Sheet.getCellByPosition(0,3).string & chr$(13) & Doc.Sheets.getByName("éphémérides").getCellByPosition(22, 2).string 
	'ajout "phase 1" au signe du soleil
	Sheet.getCellByPosition(1,2).string=Sheet.getCellByPosition(1,2).string & chr$(13) &  "lp1"



'ajoute NS au tableau progressé
Sheet = Doc.Sheets.getByName("psychogenèse")
	'lignes
	for i=3 to 111 '54 to 162
	'lecture cellule NN
	abc=Sheet.getCellByPosition(11,i).string
	'signe
	'attention chr$(13) est transformé en chr$(10) !
	car=instr(1,abc,chr$(10))
		'permet de sauter les lignes bleues intermédiaires
		if car then
			bcd=mid$(abc,1,car-1)
				'recherche du numéro de signe
				for j=0 to 11
					'cde : signe de NS=NN+6
					if signe(j)=bcd then cde=signe((j+6) mod 12) : exit for
				next j
			'ajoute degrés minutes NS=NN
				car1=instr(car+1,abc,chr$(10))
				cde=cde & chr$(10) & mid$(abc,car+1,car1-car-2)
			'ajoute lpp NS=NN+6
				car2=(val(mid$(abc,car1+4))+6) mod 12
				'12 mod 12 = 0
				if car2=0 then car2=12
				cde=cde & chr$(10) & "lpp" & str(car2)
			'ajoute lpn NS=NN+6
				car1=instr(car1+1,abc,chr$(10))
				car2=(val(mid$(abc,car1+4))+6) mod 12
				'12 mod 12 = 0
				if car2=0 then car2=12
				cde=cde & chr$(10) & "lpp" & str(car2)
			'écriture
			Sheet.getCellByPosition(13,i).string=cde
		endif
	next i
	'couleur jaune clair pour la cellule de naissance
	Sheet.getCellByPosition(13,3).cellBackColor = RGB(255, 255, 155)
	
	
'mise en couleur des lunaisons progressées lp1 pour chaque planète
'attention, ne pas déplacer après le calcul des phases génériques sinon ne marche pas !
	'planète
	for i=1 to 13
		'vérifie si cellule 1ère ligne (données de naissance) contient lpp1
		abc=sheet.getcellbyposition(i,3).getstring
		car1=instr(1,abc,"lpp 1" & chr$(10))
		'ligne
		for j=4 to 111 '55 to 162
			'lecture ligne
			cell=sheet.getcellbyposition(i,j)
			abc=sheet.getcellbyposition(i,j).getstring
			'vérifie si cellule contient lpp1
			car2=instr(1,abc,"lpp 1" & chr$(10))
			if car2 then
				'si 1ere ligne = lpp1, la 2ème ligne et les quelques lignes suivantes ne doivent pas être mises en jaune (saut de 25 lignes)
				if car1 and j=4 then
				j=j+25
				else
				'mise en couleur jaune clair
				cell.cellBackColor = RGB(255, 255, 155)
				'saute 25 lignes pour éviter d'avoir plusieurs lignes successives en jaune
				j=j+25
				endif
			endif
		next j
	next i
	
	

	
'''''''phases génériques de 1 à 90 ans, de Jupiter à Pluton''''''''''
	if form_ok=1 then feuille.getControl("Label3").text="phases génériques"
Sheet = Doc.Sheets.getByName("éphémérides")

'limites basses et hautes de la barre de progression (si form 'visuel' active)
	if form_ok=1 then
	commande1 = feuille.getControl("ProgressBar1")
	commande1.setrange(0,99)
	commande2= feuille.getControl("ProgressBar2")
'	commande2.setrange(0,final_gm)
	commande2.visible=false
	endif
	
'positions des anniversaires de 1 à 90 ans dans les éphémérides, mis dans position(1 à 90)
abc=Sheet.getCellByPosition(22, 1).getstring 'date de naissance jj/mm/aaaa
annee_naissance=val(mid$(abc,7,4)) 'aaaa
bcd=mid$(abc,1,6) 'jj/mm/
	car=17
	for i=0 to 99 '1 to 90
		'barre de progression
		if form_ok=1 then commande1.setvalue(i)
	position(i)=0
	cde=bcd & str(annee_naissance+i) 'jj/mm/aaaa+i = date anniversaire
	cde= Replace$(cde," ", "") 'suppression des espaces
			for j=car to lastrow-1
			if datevalue(Sheet.getCellByPosition(0, j).getstring)=datevalue(cde) then
			position(i)=j
			exit for
			endif
		next j
	car=j
	next i

'année de naissance hors plage éphémérides
	if position(0)=0 then
	message=" - année de naissance hors plage éphémérides, pas de calcul des phases génériques" 
	goto fin2
	endif
	
'début
'lecture positions planètes, détermination phase générique/natal, écriture
car3=0 'compteur pour saut de ligne toutes les 10 lignes

'lignes (0 à 99 ans)
	for j=0 to 99 
		'Jupiter à Pluton
		for k=5 to 9
			if position(j)=0 then exit for
	Sheet = Doc.Sheets.getByName("éphémérides")
			'position planète au jour anniversaire	
			position_ephe0=Sheet.getCellByPosition(k+1, position(j)).getvalue
			'pour la ligne de l'année de naissance, lecture des données réelles, pas celles des éphémérides
			if j=0 then position_ephe0=Sheet.getCellByPosition(k+1, 1).getvalue
			'orbe
			gap=position_ephe0 - natal(k)
			if gap<0 then gap=gap+360
			'phase
			phase=int(gap/30)+1
			bcd=chr$(966) & str(phase)
	
			'écriture feuille theme
	Sheet = Doc.Sheets.getByName("psychogenèse")
			Cell = Sheet.getCellByPosition(k+1, 3+j+int(car3/10))
			cell.string=cell.string & chr$(13) & bcd
		next k
	car3=car3+1
	next j

'correction pour Jupiter car phases = cycles de 12 ans (phase 1 = 1 à 12ans, phase 2 = 12 à 24 ans, etc.)
Sheet = Doc.Sheets.getByName("psychogenèse")
car1=4 '55 '1ère ligne
car2=1 '1ère phase pour 1 à 12 ans

 	'lignes analysées
	for i=9 to 111 '60 to 162
		abc=Sheet.getCellByPosition(6, i).getstring
		car3=instr(1,abc,chr$(966))
		'début du cycle suivant avec phase=1, donc ajout de 1 à ce cycle et mise à la même valeur des 12 lignes du cycle précédent
		if car3 and val(mid$(abc,car3+1))=1 and i > car1 or i=car1+13 then 'i > car1 pour éviter si 2 lignes successives avec phase =1, i=car1+13 s'il manque une phase 1
			'ajout de 1 à la phase du début de cycle suivant 
			Sheet.getCellByPosition(6, i).string=  mid$(abc,1,car3) & str(car2+1)
			'mise à la même valeur des phases du cycle précédent
			for j=car1 to i-1
				bcd=Sheet.getCellByPosition(6, j).getstring
				car3=instr(1,bcd,chr$(966))
				'évite la ligne avec le nom de la planète
				if car3 then
					cde=mid$(bcd,1,car3) & str(car2)
					Sheet.getCellByPosition(6, j).string=cde
				endif
			next j
		car1=i+1 'cycle suivant
		car2=car2+1 'phase=phase + 1 pour le prochain cycle
		endif
	next i
	
  'correction des dernières lignes du tableau qui n'ont pas été prises en compte (de car1 à 150)
  for i = car1 to 111 '162
	abc=Sheet.getCellByPosition(6, i).getstring
	car3=instr(1,abc,chr$(966))
	Sheet.getCellByPosition(6, i).string = mid$(abc,1,car3) & str(car2)
  next i
  
    
			
''''''''''récupération du nom de fichier .txt''''''''''
fin2:
	
'effacement éventuel cadre tableau progressé
'	if message <> "" then

'	Cell = Sheet.getCellrangeByPosition(0,1,14,112) 
'	cell.clearcontents (1 or 2 or 4 or 32) '1 : valeurs numériques, 2 : date, 4 : string, 32: formatage (dont couleur et avec recalcul du lastrow)
'	Sheet.getCellByPosition(7, 0).string=""
'	endif

'supprimme les 51 1ères lignes
'Sheet.Rows.removeByIndex(0, 51)

Sheet = Doc.Sheets.getByName("psychogenèse")

'alignement du texte : centre H et V
	cell=Sheet.getCellRangeByPosition(0,1,14,112) 
	Cell.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	Cell.VertJustify = com.sun.star.table.CellVertJustify.CENTER
	
'ajustement de la largeur des colonnes feuille psychogenèse
	for i=0 to 14
	Sheet.columns(i).Optimalwidth = True
	next i
	
'ajustement de la hauteur des lignes
	for i=2 to 112 
	Sheet.Rows(i).OptimalHeight = True
	next i


fin:
'ajustement de la largeur des colonnes feuille thème
Sheet = Doc.Sheets.getByName("thème")
	for i=0 to 17
	Sheet.columns(i).Optimalwidth = True
	next i
	
'mise au 1er plan de la feuille
	Controller = Doc.CurrentController
	controller.setActiveSheet(sheet)

'active boutons tableaux, thème, aspects, etc.
call actions_boutons2(true)
	
if message <> "" then msgbox "terminé" & chr$(10) & message
End Sub


Sub calc_deg_min(cell,offset)

	longitude=cell.getvalue+offset
	if longitude >=360 then longitude=longitude-360
	if longitude <0 then longitude=longitude+360
	'signe
	index_signe=int(longitude/30) 
	'degrés/minutes
	degres=int(longitude) mod 30
	coeff1=int(longitude/30)*30
	minutes=int(60*(longitude-coeff1-degres))
			
		'ancienne méthode (aussi rapide)
		'abc=str(longitude)
		'conversion degrés minutes
		'degres=int(longitude-(30*index_signe)) 'degrés
		'bcd=mid$(abc,instr(1,abc,".")+1,2)
		'minutes=int(val(bcd)*6/10) 'minutes
				
End Sub


REM  *****  BASIC  *****


Sub ephemerides
dim ligne as long 'integer = 32768 lignes max
dim nom(1 to 100) as string
dim choixnom as integer
dim nomfichier as string
dim anneemin, anneemax as string
dim rangmin, rangmax as integer
dim bm(0 to 30000,0 to 12) as double 'colonnes date + planètes éphémérides

goto debut1

'********obsolète******************************
'vérifie présence fichier de type Ephe...txt
	abc=Dir(curdir & "/Ephe*.txt")
	if abc = "" then msgbox "pas de fichier 'Ephemerides.txt' trouvé !" : exit sub
	
'liste des fichiers .txt
	nom(1)=abc : car= 2
	do until abc=""
	abc=dir()
	if abc <> "" then nom(car)=abc : car=car+1
	loop
	nommax=car-1

'écrit liste dans abc
	abc=""
	for i=1 to nommax
	abc=abc & str(i) & "=" &  nom(i) &","'& chr$(13)
	next i

'choix du fichier .txt
question:
	choixnom=inputbox (abc, "choix du fichier d'éphémérides","1")
	if choixnom=0 then exit sub
	if choixnom > nommax then msgbox nommax & " max !" : goto question
	nomfichier=nom(choixnom)
	
'********fin obsolète******************************

debut1:	
'ouvre sélecteur de fichiers
	abc=""
	call FileButtonSelected
	if abc="" then exit sub
	nomfichier=abc 
	
'détermination dates de début et de fin du fichier éphémérides
	open nomfichier for input as #1
	anneemin=""
		Do While Not EOF(1)
		
		 Line Input #1, abc
		    if val(abc) and anneemin="" then
			    bcd=mid$(abc,1,10)
			    anneemin = mid$(abc,7,4)
			    'positionnement à la fin du fichier - 200 bytes
			    seek(1,lof(1)-200)
			endif
		Loop
	close #1
	cde= mid$(abc,1,10)
	anneemax = mid$(abc,7,4)

if anneemin =""  or anneemax="" then msgbox "ce n'est apparemment pas un fichier éphémérides !" : goto debut1
'choix des dates de début et fin
ref1:
	choix=inputbox ("entrer année de départ" , "éphémérides " & bcd & " - " & cde , anneemin)
	if choix="" then exit sub
	if choix < anneemin or choix > anneemax then goto ref1
	anneemin = choix
ref2:
	choix=inputbox ("entrer année de fin" , "éphémérides " & bcd & " - " & cde , anneemax)
	if choix="" then exit sub
	if choix < anneemin or choix > anneemax then goto ref2
	anneemax = choix

'confirmation
	choix=inputbox ("confirmer chargement du fichier  " & nomfichier & "  dans la feuille éphémérides ? (effacement à partir de la ligne 18) (O, N) ", "éphémérides " & anneemin & " - " & anneemax,"O")
	if choix <> "O" then exit sub

Doc = ThisComponent

'désactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(false)

'vérifie présence feuille éphémérides, sinon création
If not Doc.Sheets.hasByName("éphémérides") Then 
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
	Doc.Sheets.insertByName("éphémérides", Sheet)
else
	'récupère le nombre de lignes utilisées dans la feuille éphémérides
	Sheet = Doc.Sheets.getByName("éphémérides")
	Curs = Sheet.createCursor
	Curs.gotoEndOfUsedArea(True)
	lastrow = Curs.Rows.Count 'ne pas mettre -1 car erreur s'il n'y a pas d'éphémérides (lastrow=17)
	'effacement feuille éphémérides (sauf la partie haute du thème (lignes 0 à 17) atttention clearcontents(32) remet la taille des caractères à 10
	Cell = Sheet.getCellrangeByPosition(0,17,40,lastrow) 
	cell.clearcontents (1 or 2 or 4 or 32) '1 : valeurs numériques, 2 : date, 4 : string, 32: formatage (dont couleur et avec recalcul du lastrow)
endif

sheet.charheight=7
'date de début
	Sheet.getCellByPosition(13, 17).string="01/01/" & anneemin
'ajustement de la largeur des colonnes soit /1ère ligne soit/ligne 18
	if Sheet.getCellByPosition(1, 0).string="" then
		Sheet.getCellByPosition(0, 17).string="00/00/0000"
		for i=1 to 12
		Sheet.getCellByPosition(i, 17).string="Mercure "
		next i
		Sheet.getCellByPosition(13, 17).string="00:00:0000"
		Sheet.getCellByPosition(14, 17).string="0000 - 0000"
	endif
	for i=0 to 20
	Sheet.columns(i).Optimalwidth = True
	next i

'format date pour le compteur de date
	Sheet.getCellByPosition(13, 17).numberformat=10030
'réinitialise l'array à sa taille maximum sinon pb au 2ème lacement du programme si taille array > précédente
	redim bm(0 to 30000,0 to 12)

'limites basses et hautes de la barre de progression (si form 'visuel' active)
	if form_ok=1 then
	feuille.getControl("ProgressBar1").setrange(val(anneemin),val(anneemax))
	feuille.getControl("ProgressBar2").visible=false
	'initialise la barre de progression
	feuille.getControl("ProgressBar1").setvalue(0)
	feuille.getControl("Label3").text=""
	endif
	
'début
	'lecture fichier "Ephemerides.txt" et remplissage feuille à partir de la ligne 18 
	open nomfichier for input as #1
	'compteur de lignes dans arrays a() et gm()
	ligne=0
	'compteur de lignes écrites dans éphémérides
	compte_lignes=17
car=lof(1)	
on error goto fin
Do While Not EOF(1)
 'lecture 1 ligne
 Line Input #1, abc
 'vérifie si ligne contient des positions de planètes
 if val(abc) then
 	'année
 	bcd=mid$(abc,7,4)
	'année hors plage,sortie
	if val(bcd) > anneemax then exit do
   	'ok
   	'plus lent !if val(bcd) >= val(anneemin) and val(bcd) <= val(anneemax) then
   	if val(bcd) >= val(anneemin) then if val(bcd) <= val(anneemax) then
   	  'écriture d'une ligne
      	  'date
   	  	  bm(ligne,0)=datevalue(mid$(abc,1,10))
   		  'positions Soleil à Lilith dans array bm(x,y)
		   car= 22 '22 = position Soleil
		   		for i = 1 to 12
		   		bm(ligne,i)=val(mid$(abc,car,9))
		   		car=car+ 11 
		     	next i
		     		'écriture dans éphémérides d'un bloc de 30000 lignes (doit correspondre à definition de bm(0 to 30000)
					if ligne =30000 then
					call ecrit_ephemerides(ligne,compte_lignes,bm)
					'réinitialise ligne
					ligne=-1
					endif
	  ligne=ligne+1 
	endif
  endif
Loop
close #1

'écrit dernier bloc de lignes
ligne=ligne-1
call ecrit_ephemerides(ligne,compte_lignes,bm)

'format date pour la 1ere colonne
Cell = Sheet.getCellrangeByPosition(0,17,0,compte_lignes) 
cell.numberformat=10030

'pas de décimales pour les longitudes
Cell = Sheet.getCellrangeByPosition(1,17,12,compte_lignes) 
cell.numberformat=10001 ' pas de décimales

'2 décimales pour l'orbe du graphe
Cell = Sheet.getCellByPosition(0,0) 
cell.numberformat=10002

'affichage de la plage de dates utilisée
Cell = Sheet.getCellByPosition(14, 17) : cell.string=anneemin & " - " & anneemax
Cell.CellBackColor = RGB(0, 255, 0)
Cell = Sheet.getCellByPosition(13, 17)
Cell.CellBackColor = RGB(0, 255, 0)

'écrit "non" cellule Q18
	Sheet.getCellByPosition(16, 17).string = "non"
	
'alignement au centre
sheet.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER

'ajustement largeur des colonnes N et O
for i=13 to 14
Sheet.columns(i).Optimalwidth = True
next i

'affichage du nouveau nom dans la form (si form 'visuel' active)
	if form_ok=1 then
	commande1 = feuille.getControl("Label1")
	commande1.text=Sheet.getcellbyposition(14,17).getstring
	'phases ? (non)
	feuille.getControl("Label12").text="non"
	endif
	
'mise au 1er plan de la feuille
Controller = Doc.CurrentController
controller.setActiveSheet(sheet)

'sactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(true)

msgbox "terminé, ligne " & str(compte_lignes) & chr$(13) & message : exit sub

fin:
message = message + chr$(13) + "attention, erreur " & mid$(abc,1,10) & " - ligne : " & str(ligne+1) 
resume next
End Sub


Sub ecrit_ephemerides(ligne,compte_lignes,bm)
	'affiche la date de fin du bloc
	cell= Sheet.getCellByPosition(13, 17)
	cell.value=bm(ligne,0)
	'incrémente barre de progression
	if form_ok=1 then feuille.getControl("ProgressBar1").setvalue(year(cell.value))
	'efface lignes superflues  
	redim preserve bm(0 to ligne, 0 to 12)
				'arrondi à 3 décimales dans array bb (pose des problèmes avec la détermination des phases !)
				'bb=oFA.callFunction("Round",array(bm,3))
	'écriture données
	Cell = Sheet.getCellrangeByPosition(0,compte_lignes,12,compte_lignes+ligne)
	cell.setData(bm) 'setdata pour nombres seulement (marche aussi avec setdataarray)
				'cell.setData(bb)
	compte_lignes=compte_lignes+ligne+1
end sub


sub phases_ephemerides
dim phase1_val, phase2_val
dim phase0_rang as long, phase1_rang as long, phase2_rang as long,phase3_rang as long
dim compteur as long
dim debut as long,final as long
dim dk' (17 to 100000,3 to 10) as double array pour calcul des phases éphémérides
dim final_dk as long
dim phase_ephe(0 to 300,0) as string 'our les phases 1,2,3 - attention, laisser 2 colonnes (x,0) sinon pas d'écriture dans la feuille avec setdata !

Doc = ThisComponent

'vérifie présence feuille éphémérides, sinon création
	If not Doc.Sheets.hasByName("éphémérides") Then msgbox "pas de feuille éphémérides !, exécuter 'ephemerides' d'abord" : exit sub

Sheet = Doc.Sheets.getByName("éphémérides")
	if Sheet.getCellByPosition(0, 17).getstring ="" then msgbox "le tableau des éphémérides est vide, exécuter 'a1bis_ephemerides' d'abord" : exit sub

'conirmation
	choix=inputbox ("calcul et mise en couleurs des phases feuille éphémérides (O,N) ?", "couleurs phases","O")
	if choix <> "O" then exit sub

'désactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(false)

'remise à 0 des couleurs (900000 lignes)
	Cell = Sheet.getCellrangeByPosition(3,17,10,900000) : Cell.CellBackColor = RGB(255, 255, 255)
'format de date pour compteur date
	Cell = Sheet.getCellByPosition(14, 17) : cell.numberformat=10030 
 
'récupère le nombre de lignes utilisées dans la feuille éphémérides
	Sheet = Doc.Sheets.getByName("éphémérides")
	Curs = Sheet.createCursor
	Curs.gotoEndOfUsedArea(True)
	lastrow = Curs.Rows.Count-1

'effacement des colonnes contenant les phases (attention clearcontents(32) remet la taille des caractères à 10)
	Cell = Sheet.getCellrangeByPosition(23,17,30,lastrow) 
	cell.clearcontents (1 or 2 or 4) '1 : valeurs numériques, 2 : date, 4 : string, 32: formatage (dont couleur et avec recalcul du lastrow)
	
'array dk() a un maximum de 16368 lignes donc découpage en tronçons de 16000 lignes
'debut et fin de la plage de lecture éphémérides
debut=17 : final=16017
if final >= lastrow then final=lastrow
'fin de la plage d'écriture dans array dk() début à 0
final_dk=final-17 '16000 max
if final_dk >= 16000 then final_dk=16000

'on error goto fin

'limites basses et hautes de la barre de progression (si form 'visuel' active)
	if form_ok=1 then
		commande1=feuille.getControl("ProgressBar1")
		commande1.setrange(0,int(lastrow/16017))
		if lastrow < 16017 then commande1.setrange(0,1)
		commande2=feuille.getControl("ProgressBar2")
		commande2.setrange(-1,7)
		commande2.visible=true
		feuille.getControl("Label3").text=""
	endif


'début 
for k= 0 to int(lastrow/16017)
	if form_ok=1 then commande1.setvalue(k)
	'affichage date du bloc en cours
	Sheet.getCellByPosition(13, 17).string = Sheet.getCellByPosition(0, debut).getstring 
	'écriture plage de données dans array dk() (16000 lignes)
'	Cell = Sheet.getCellrangeByPosition(3,debut,10,final)
	Cell = Sheet.getCellrangeByPosition(3,debut,12,final)
	dk()=cell.getData

'recherche des phases dans array dk()
	'de Mercure à Lilith 'Pluton
	for i = 0 to 9 '7
	'saute NN
	if i=8 then goto fini
	'affiche progression curseur
	if form_ok=1 then commande2.setvalue(i)
	'affichage planète en cours d'analyse
	Sheet.getCellByPosition(15, 17).string = Sheet.getCellByPosition(i+3, 0).getstring 
		'compteur de lignes dans array dk()
		compteur=1
		'sortie si car=-1
		car=0 
		'recherche phase 1 (max)
		do '0
		if compteur<= 0 or compteur >= final_dk then exit do
		call calcul_trois_lignes_array(i,compteur,dk)	
		 'phase 1 trouvée
		 'vitesse idem ! if position_ephe2 >= position_ephe1 and position_ephe2 >= position_ephe3 then
		 if position_ephe2 >= position_ephe1 then if position_ephe2 > position_ephe3 then
		 phase1_rang = compteur
		 phase1_val = position_epheref2
		 car=0
			'recherche phase 2 (min)
			do '1
			if compteur >= final_dk then car=-1 : exit do
			call calcul_trois_lignes_array(i,compteur,dk)
				'phase 2 trouvée
				'vitesse idem ! if position_ephe2 <= position_ephe1 and position_ephe2 <= position_ephe3 then
				if position_ephe2 <= position_ephe1 then if position_ephe2 < position_ephe3 then
				phase2_rang = compteur : phase2_val = position_epheref2
					'recherche phase 0 (=début phase 1)
					compteur=phase1_rang
					do '2
					'très important sinon boucle sans fin !
					if compteur > 0 then call calcul_trois_lignes_array(i,compteur,dk)
							'si on atteint le début du fichier, phase 0 tronquée
							if compteur <= 0 then
							phase0_rang = 0 
							'phase0_val = Sheet.getCellByposition(i+3, 17).getvalue
							position_ephe3=phase2_val + 1 ' pour que le test "if" suivant soit ok
						 	position_ephe1=phase2_val - 1 'idem
						 	endif
						'phase 0 trouvée (2ème test au cas où phase2_val proche de 0 et les 2 valeurs à comparer ont été incrémentées de 360)
						if phase2_val <= position_ephe3 and phase2_val > position_ephe1 or phase2_val+360 <= position_ephe3 and phase2_val+360 > position_ephe1 then
						phase0_rang = compteur
							'recherche phase 3
							compteur=phase2_rang
							do '3
							'if compteur<= debut_bm or compteur >= final_bm then car=-1 : exit do
							if compteur < final_dk then call calcul_trois_lignes_array(i,compteur,dk)
									'si on atteint la fin du fichier, phase 3 tronquée
									if compteur >= final_dk then
									phase3_rang = compteur
									position_ephe3=phase1_val + 1 ' pour que le test "if" suivant soit ok
						 			position_ephe1=phase1_val - 1 'idem
									endif
								'phase 3 trouvée (2ème test au cas où phase1_val proche de 0 et les 2 valeurs à comparer ont été incrémentées de 360)
								if (phase1_val <= position_ephe3 and phase1_val > position_ephe1) or (phase1_val+360 <= position_ephe3 and phase1_val+360 > position_ephe1) then
								phase3_rang = compteur
								
					'******mise en couleur (vert,jaune, rouge) et ajout des phases 20 colonnes plus loin (de X à AE)******
					call ecrit_phases(i+3,0,phase0_rang+debut,phase1_rang+debut,phase0_rang,phase1_rang,"phase 1",RGB(167, 236, 106))
					call ecrit_phases(i+3,1,phase1_rang+debut,phase2_rang+debut,phase1_rang,phase2_rang,"phase 2",RGB(255, 213, 6))
					call ecrit_phases(i+3,1,phase2_rang+debut,phase3_rang+debut,phase2_rang,phase3_rang,"phase 3",RGB(255, 153, 153))
							
									'sortie
									compteur=phase3_rang
									car=-1
									exit do
								endif 'phase 3
							if car=-1 then exit do
							compteur=compteur+1
							loop '3
						exit do
						endif 'phase 0
					if car=-1 then exit do
					compteur=compteur-1
					loop '2
				endif 'phase 2
			if car=-1 then exit do
			compteur=compteur+1
			loop '1
		 endif'phase 1
		compteur=compteur+1
		loop '0
fini:
	next i
	
	'incrément de la barre de progression
	if form_ok=1 then commande1.setvalue(k+1)
'incrémentation des plages de lecture/écriture
	'début
	debut=debut+16000
	if debut > lastrow then exit for
	'fin
	final=final+16000
	'si fin de fichier, dernier tronçon plus court
	if final > lastrow then
	final_dk=16000-(final-lastrow)
	final=lastrow
	 endif
		
next k

'affichage date de fin
	Sheet.getCellByPosition(13, 17).string = Sheet.getCellByPosition(0, lastrow).getstring
	 
'écrit "oui" cellule Q18
	Sheet.getCellByPosition(16, 17).string = "oui"
	
'ajustement de la largeur des colonnes
	for i=13 to 14
	Sheet.columns(i).Optimalwidth = True
	next i
'mise au 1er plan de la feuille
	Controller = Doc.CurrentController
	controller.setActiveSheet(sheet)

'mis phases=oui (si form 'visuel' active)
	if form_ok=1 then
	feuille.getControl("Label12").text="oui"
	endif
	
'sactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(true)
	
msgbox "terminé"
exit sub

fin:
msgbox "erreur planète (i= " & str(i) & " ) " & planete(i+3) & " ligne : " & str(val(debut)+val(compteur))

end sub

sub calcul_trois_lignes_array(i as integer, compteur as long,dk)

'3 longitudes : ligne précédente, en cours et suivante
	position_ephe1=dk(compteur-1)(i)
	position_ephe2=dk(compteur)(i)
	position_ephe3=dk(compteur+1)(i)

'sauvegade longitude de la ligne en cours
	position_epheref2=position_ephe2
	
'ajoute 360 si passage de 359 à 0
	if position_ephe1 - position_ephe2 > 300 then position_ephe2 = position_ephe2 + 360 : position_ephe3=position_ephe3+360
	if position_ephe2 - position_ephe3 > 300 then position_ephe3 = position_ephe3 + 360
	if position_ephe2 - position_ephe1 > 300 then position_ephe1 = position_ephe1 + 360
	if position_ephe3 - position_ephe2 > 300 then position_ephe1 = position_ephe1 + 360 : position_ephe2 = position_ephe2 + 360

end sub


Sub ecrit_phases(col,inc,lig1,lig2,rang1,rang2,texte,teinte)
dim phase_ephe(0 to 300,0) as string  'attention, laisser 2 colonnes (x,0) sinon pas d'écriture dans la feuille avec setdata 

	'mise en couleur
	Sheet.getCellRangeByPosition(col,lig1+inc,col,lig2).CellBackColor = teinte
	'inutile (gardé car ne change pas le temps d'exécution) ? réinitialise array à sa taille maximum sinon pb au 2ème lacement du programme si taille array > précédente
	redim phase_ephe(0 to 300,0)
	'écriture phase dans array phase_ephe() (inc=0 pour phase 1 et inc=1 pour phases 2 et 3)
	for j=0 to rang2 - rang1 -inc
	phase_ephe(j,0)=texte
	next j
	'redimensionnement array à la taille utile
	redim preserve phase_ephe(0 to rang2 - rang1-inc,0)
	'écriture phase dans la feuille
	cell=Sheet.getCellRangeByPosition(col+20,lig1+inc,col+20,lig2)
	cell.setDataArray(phase_ephe())
										
End Sub


REM  *****  BASIC  *****


REM  *****  BASIC  *****


Sub tableau_transits_old
dim jours as long
dim compteur as long
dim date_transit as string
dim transitante as string
Dim position(0 to 256) as double
dim position1, position2 as double
dim orbe as double
dim signemoins as string
dim datedebut, datefin as string
dim anneemin as string
dim anneemax as string
dim an(0 to 200) as string
dim rangmin(0 to 200) as long
dim rangmax(0 to 200) as long
dim rangminimum as long
dim rangmaximum as long
dim couleur_planete as long
dim tabindex as integer
dim tableau as integer
dim nom_tableau as string
dim phase_generique as string

Doc = ThisComponent
'vérifie si feuille éphémérides présnte
If not Doc.Sheets.hasByName("éphémérides") Then msgbox "pas de feuille éphémérides !, abandon" : exit sub

Sheet = Doc.Sheets.getByName("éphémérides")
' lit l'orbe transit pour chaque aspect
for i = 1 to 16
orbe= Sheet.getCellByPosition(18, i).getvalue
if orbe <0 or orbe > 10 then msgbox "feuille 1 colonne H ligne"& str(i+1) &", la valeur d'orbe : "  & orbe &" est incorrecte (<0 ou > 10)"   : exit sub
next i
' vérifie si données thème et éphémérides sont présentes
if Sheet.getCellByPosition(1, 1).getstring ="" then  msgbox "le thème est vide, exécuter 'calcul_du_thème' d'abord" : exit sub
if Sheet.getCellByPosition(0, 17).getstring ="" then  msgbox "le tableau des éphémérides est vide, exécuter 'ephemerides' d'abord" : exit sub

'choix planètes
ref0:
car=val(inputbox ("création du tableau de transits des planètes personnelles (1) ou mondiales (Jupiter à Pluton) ? (2)" , "tableau transits", 2))
if car=0 then exit sub
if car < 1 or car > 2 then goto ref0
if car=1 then tableau=1 : num0=1 : num1= 5 : nom_tableau="transits1" else tableau=2 : num0=6 : num1=12 : nom_tableau="transits2"

'définitions variables aspect et planete du theme
call definitions
Sheet = Doc.Sheets.getByName("éphémérides")
'détermination années min et max des éphémérides avec les lignes de début et fin pour chaque année
compteur =17
an(0) = right (Sheet.getCellByPosition(0, 17).getstring,4)
anneemin=an(0)
rangmin(0)=17 : i=0
	do until Sheet.getCellByPosition(0, compteur).getstring =""
	abc = right (Sheet.getCellByPosition(0, compteur).getstring,4) 'année
		if abc <> an(i) then 
		rangmax(i)=compteur-1
		i =i+1
		rangmin(i)=compteur
		an(i)=abc
		compteur=compteur+360 'saut d'une année, ok seulement si on démarre du 1er janvier !
		endif
	compteur=compteur+1
	loop
anneemax=abc
rangmax(i)=compteur	-1 'dernière ligne dernière année


'choix des années de début et fin
ref1:
choix=inputbox ("entrer année de départ" ,"tableau de transits " & anneemin & " - " & anneemax, an(0))
if choix="" then exit sub
if choix < anneemin or choix > anneemax then goto ref1
anneemin = choix
car=val(anneemin)-val(an(0))
rangminimum=rangmin(car) '1ere ligne à utiliser
ref2:
choix=inputbox ("entrer année de fin" ,"tableau de transits " & anneemin & " - " & anneemax, anneemax)
if choix="" then exit sub
if choix < anneemin or choix > anneemax then goto ref2
anneemax = choix
car=val(anneemax) - val(an(0))
rangmaximum=rangmax(car) 'dernière ligne à utiliser

'confirmation
choix=inputbox ("confirmer effacement feuille " & nom_tableau & " et création du tableau des transits de " & anneemin & " à " & anneemax & " (O,N) ?", "tableau des transits","O")
if choix <> "O" then exit sub

'suppression et recréation feuille transits
	If Doc.Sheets.hasByName(nom_tableau) Then  Doc.Sheets.RemoveByName(nom_tableau)'suppression feuille
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
	Doc.Sheets.insertByName(nom_tableau, Sheet)

'récupération du nom du fichier de thème utilisé (ex. astro.txt)
	Sheet = Doc.Sheets.getByName("éphémérides")
	abc=sheet.getCellByPosition(19, 0).getstring
	Sheet = Doc.Sheets.getByName(nom_tableau)
	Cell = Sheet.getCellByPosition(7,0) : cell.string= abc
	Cell.CellBackColor = RGB(0, 255, 0)
	'fomat de date = jj/mm/aaaa pour le compteur
	Cell = Sheet.getCellByPosition(6,0)
	cell.numberformat=10030 
	Cell.CellBackColor = RGB(0, 255, 0)
	sheet.charheight=7

'récupération des orbes transits max et de toutes les positions du thème de position(0) à position(239)
Sheet = Doc.Sheets.getByName("éphémérides")
car=0
for i=1 to 16 'aspects
orbe_transit(i-1)=Sheet.getCellByPosition(18, i).getvalue 'lecture des orbes feuille éphémérides plutôt que dans "definitions"
if tableau = 1 then orbe_transit(i-1)=orbe_transit(i-1) * 10 'pour les personnelles, orbes multipliés par 10
for j=1 to 16 'planetes
position(car)= Sheet.getCellByPosition(j, i).getvalue
car=car+1
next j
next i

'écriture en-tête feuille 3
	Sheet = Doc.Sheets.getByName(nom_tableau)
	Sheet.getCellByPosition(0,0).string= "date"
	Sheet.getCellByPosition(1,0).string= "transit de"
	Sheet.getCellByPosition(2,0).string= " en position"
	Sheet.getCellByPosition(3,0).string= "en aspect"
	Sheet.getCellByPosition(4,0).string= "sur planète"
	Sheet.getCellByPosition(5,0).string= "orbe transit"
	Sheet.getCellByPosition(6,0).numberformat=10030 'format date

'début
Sheet = Doc.Sheets.getByName("éphémérides")
compte_lignes=1 'compteur ligne feuille transit

'compteur lignes feuille épémérides
for compteur=rangminimum to rangmaximum

'récupération des positions des planètes d'une journée et comparaison avec toutes les positions du theme
 for i = num0 to num1 'positions planètes d'une journée
	 longitude= Sheet.getCellByPosition(i, compteur).getvalue
 	'finalement n'utilise que le 1er orbe (orbe_transit (0)) pour accélérer car tous les aspects ont le même orbe
	 position1 = longitude-orbe_transit(0) : position2 = longitude+orbe_transit(0) 'tolérance orbe du transit
				
	'positions du thème
	for j=0 to 255 
      'transit trouvé
	 'plus lent ! if position(j) >= position1 and position(j) <= position2 then
	  if position(j) >= position1 then if position(j) <= position2 then
	  	date_transit=Sheet.getCellByPosition(0, compteur).getstring
 		transitante=planete(i-1)
		couleur_planete=Sheet.getCellByPosition(i,compteur).cellbackcolor
		'position transit en degrés minutes + signe
		coeff1=int(longitude/30)*30 : coeff2=int(longitude) mod 30
		abc=str(coeff2) & chr$(176) & str(int(60*(longitude-coeff1-coeff2))) & "' " & signe(coeff1/30) 'ex  12° 23' Lion
		'orbe transit en degrés minutes
		orbedecimal= longitude - position(j)
		signemoins="+"
		if orbedecimal < 0 then orbedecimal=abs(orbedecimal) : signemoins="-"
		coeff1=int(orbedecimal/30)*30 : coeff2=int(orbedecimal) mod 30
		bcd=signemoins & str(coeff2) & chr$(176) & str(int(60*(orbedecimal-coeff1-coeff2))) & "'"
		
						'calcul phase générique si planète collective (Jupiter - Pluton) ; pas utile à priori !
					'	phase_generique=""
					'	if nom_tableau="transits2" then
							'longitude du jour-longitude de naissance 
					'		position1=longitude-position(i-1)
					'		'phase décroissante ?
					'		if position1<0 then position1=position1+360
							'phase progressée
					'		position2=int(position1/30)+1
					'		phase_generique=chr$(966) & str(position2)
					'	endif
						
		'écriture ligne dans tableau transits
		Sheet = Doc.Sheets.getByName(nom_tableau)
			'compteur de date
			Sheet.getCellByPosition(6, 0).string = date_transit
			'date au format numérique
			Sheet.getCellByPosition(0, compte_lignes).string = date_transit
			'planète du jour
			Cell = Sheet.getCellByPosition(1, compte_lignes) : cell.string= transitante
			Cell.CellBackColor =couleur_planete
			'position degrés + signe
			Sheet.getCellByPosition(2, compte_lignes).string= abc
			'aspect = aspect(valeur entière de j/10)
			Sheet.getCellByPosition(3, compte_lignes).string= aspect(int(j/16))
			'planète du thème
			Sheet.getCellByPosition(4, compte_lignes).string= planete(j mod 16)
			'orbe
			Sheet.getCellByPosition(5, compte_lignes).string= bcd
			'phase
			Cell = Sheet.getCellByPosition(6, compte_lignes)
				select case couleur_planete
				case RGB(167, 236, 106) 'vert
				cell.string= "phase 1"
				case RGB(255, 213, 6)  'orange
				cell.string= "phase 2 (R)"
				case RGB(255, 153, 153) 'rouge
				cell.string= "phase 3"		 
				end select
					'Sheet.getCellByPosition(7, compte_lignes).string=phase_generique 'colonne H
		compte_lignes= compte_lignes+1
		'permet de sauter les aspects restants de la même planète
		j=j+15-(j mod 16) 
		Sheet = Doc.Sheets.getByName("éphémérides")
	  endif
	next j
 next i
next compteur

'changement formatage
Sheet = Doc.Sheets.getByName(nom_tableau)
Cell = Sheet.getCellrangeByPosition(0,1,0,compte_lignes) ': cell.numberformat=10030 'date jj/mm/aaaa sur 1ère colonne, lignes 0 à compte_lignes
sheet.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER

'ajustement de la largeur des colonnes
	for i=0 to 7
	Sheet.columns(i).Optimalwidth = True
	next i

'mise au 1er plan de la feuille
Sheet = Doc.Sheets.getByName(nom_tableau)
Controller = Doc.CurrentController
controller.setActiveSheet(sheet)
msgbox "terminé : " & compte_lignes & " lignes écrites"
End sub


sub cadre_cellules(feuille as string, colfin as integer,ligfin as integer)
Dim aBorder as New com.sun.star.table.BorderLine
Dim oBorder as New com.sun.star.table.TableBorder

Doc=thiscomponent
Sheet = Doc.Sheets.getByName(feuille)
	'lignes bleues autour des cellules
	aBorder.Color = RGB(0,204,204)
	aBorder.InnerLineWidth = 0
	aBorder.OuterLineWidth = 10
	aBorder.LineDistance = 0
	
	oBorder.LeftLine = aBorder
	oBorder.TopLine = aBorder
	oBorder.RightLine =aBorder
	oBorder.BottomLine = aBorder
	oBorder.VerticalLine =aBorder
	oBorder.HorizontalLine =aBorder
	
	oBorder.isLeftLineValid =true 'false par défaut
	oBorder.isTopLineValid =true 'false par défaut
	oBorder.isRightLineValid =true 'false par défaut
	oBorder.isBottomLineValid =true 'false par défaut
	oBorder.isVerticalLineValid =true 'false par défaut
	oBorder.isHorizontalLineValid =true 'false par défaut
	'tracé des lignes
	sheet.getCellRangeByPosition(0,0,colfin,ligfin).TableBorder = oBorder
end sub



Sub tableau_transits ' pas de sous-routine
Dim aSortFields(0) As New com.sun.star.util.SortField 'pour trier colonnes
Dim aSortDesc(0) As New com.sun.star.beans.PropertyValue 'pour trier colonnes
'variables
dim annee1, annee2
dim anneemax as string
dim anneemin as string
dim choix1 as integer, choix2 as integer
dim choix_pro_natal
dim compteur as long
dim coul as double
dim couleur_planete as long
dim datedebut, datefin as string
dim debut as long
dim final as long
dim final_gm as long
dim gap as double
dim jours as long
dim nom_tableau as string
dim phase_generique as string
dim rangmaximum as long
dim rangminimum as long
dim signemoins as string
dim tabindex as integer
dim tableau as integer
dim val1 as integer
'arrays
dim aa 'array pour copie du thème
dim an(0 to 200) as string

dim col_a() 'array de lecture dates éphémérides
dim col_gm() 'array de lecture données planètes dans éphémérides
dim filtre(3) as string
dim matrice(1 to 16,1 to 16) as double
dim natal(0 to 15) as double 'longitudes thème natal
dim phase() 'array de lecture phases
Dim position(0 to 255) as double
dim rangmax(0 to 200) as long
dim rangmin(0 to 200) as long
dim tr(0 to 300000,0 to 6) 'as string 'array d'écriture dates+données des planètes en transit

'oFA = createUnoService("com.sun.star.sheet.FunctionAccess")
Doc = ThisComponent
'vérifie si feuille éphémérides présnte
	If not Doc.Sheets.hasByName("éphémérides") Then msgbox "pas de feuille éphémérides !, abandon" : exit sub

Sheet = Doc.Sheets.getByName("éphémérides")
val(anneemin)
' vérifie si données thème et éphémérides sont présentes
	if Sheet.getCellByPosition(1, 1).getstring ="" then  msgbox "le thème est vide, exécuter 'calcul_du_thème' d'abord" : exit sub
	if Sheet.getCellByPosition(0, 17).getstring ="" then  msgbox "le tableau des éphémérides est vide, exécuter 'ephemerides' d'abord" : exit sub

'choix transits mondiaux ou progressés
ref0:
	choix1=val(inputbox("création d'un tableau de transits mondiaux (1) ou progressés (2) ?", "choix transits mondiaux ou progressés", 1))
	if choix1=0 then exit sub
	if choix1 <1 or choix1 > 2 then goto ref0 

'choix planètes mondiales/personnelles ou natal/progressé	
if choix1=1 then
ref1:
	car=val(inputbox ("transits des planètes mondiales (1) ou personnelles (2)  ? " , "tableau transits mondiaux", 1))
	if car=0 then exit sub
	if car < 1 or car > 2 then goto ref1
	if car=2 then tableau=1 : num0=3 : num1= 5 : nom_tableau="transits1" else tableau=2 : num0= 6: num1=12 : nom_tableau="transits2"
else
ref2:
	choix_pro_natal=val(inputbox ("transits thème progressé/natal (1) ou progressé/progressé (2) ?", "tableau transits progressés", 1))
	if choix_pro_natal=0 then exit sub 
	if choix_pro_natal < 1  or choix_pro_natal > 2 then goto ref2
	nom_tableau="progressé"
endif

'désactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(false)
	
'détermination années min et max des éphémérides avec les lignes de début et fin pour chaque année
an(0) = year(Sheet.getCellByPosition(0, 17).getstring)
anneemin=an(0)
rangmin(0)=17 : i=0
compteur =377 '(17+360)

	do until Sheet.getCellByPosition(0, compteur).getstring =""
	'date jj/mm/aaaa
	bcd=Sheet.getCellByPosition(0, compteur).getstring
	'année aaaa
	abc = year(bcd) 
		if abc <> an(i) then 
			rangmax(i)=compteur-1
			i =i+1
			rangmin(i)=compteur
			an(i)=abc
			'saut d'une année pour accélérer, ok seulement si on démarre du 1er janvier et ne pas mettre 365 sinon décalage d'1 jour !
			compteur=compteur+364
		endif
	compteur=compteur+1
	loop
	
anneemax=abc
'dernière ligne dernière année
rangmax(i)=compteur-1

'annee min et annee max différentes selon mondiaux ou progressés
 if choix1=1 then
	 annee1=anneemin
	 annee2=anneemax
 else
	 annee1=year(now)
	 annee2=annee1
 endif

'choix des années de début et fin
ref10:
	choix=inputbox ("entrer année de départ" ,"tableau de transits " & anneemin & " - " & anneemax, annee1)
	if choix="" then exit sub
	if val(choix) < val(anneemin) or val(choix) > val(anneemax) then goto ref10
	anneemin = choix
	car=val(anneemin)-val(an(0))
	rangminimum=rangmin(car) '1ere ligne à utiliser
ref11:
	choix=inputbox ("entrer année de fin" ,"tableau de transits " & anneemin & " - " & anneemax, annee2)
	if choix="" then exit sub
	if choix < anneemin or choix > anneemax then goto ref11
	anneemax = choix
	car=val(anneemax) - val(an(0))
	rangmaximum=rangmax(car) 'dernière ligne à utiliser
	
'choix méthode si transits mondiaux
if choix1=1 then
ref12:
	choix2=inputbox ("choix méthode : (1) = 4sec/an, 1'/20 ans, 13'30sec./200 ans ou (2) = 3'30sec/an, 7'/200 ans ?", "tableau des transits","1")
	if choix2 = "" then exit sub
	if choix2 < "1" or choix2 > "2" then goto ref12
endif

'confirmation
	choix=inputbox ("confirmer effacement feuille " & nom_tableau & " et création du tableau des transits de " & anneemin & " à " & anneemax & " (O,N) ?", "tableau des transits","O")
	if choix <> "O" then exit sub

'définitions variables aspect et planete du theme
call definitions

	
'suppression et recréation feuille transits
	If Doc.Sheets.hasByName(nom_tableau) Then  Doc.Sheets.RemoveByName(nom_tableau)'suppression feuille
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
	Doc.Sheets.insertByName(nom_tableau, Sheet)

'récupére le nom de fichier du thème utilisé (ex. astro.txt)
Sheet = Doc.Sheets.getByName("éphémérides")
	abc=sheet.getCellByPosition(19, 0).getstring
	
'récupère le nombre de lignes utilisées dans la feuille éphémérides
	Curs = Sheet.createCursor
	Curs.gotoEndOfUsedArea(True)
	lastrow = Curs.Rows.Count-1
	
'écriture du nom feuille transits + couleur verte
Sheet = Doc.Sheets.getByName(nom_tableau)
	Cell = Sheet.getCellByPosition(7,0)
	cell.string= abc
	Cell.CellBackColor = RGB(0, 255, 0)
	
'écriture en-tête feuille transits
	Sheet.getCellByPosition(0,0).string= "date"
	Sheet.getCellByPosition(1,0).string= "transit de"
	Sheet.getCellByPosition(2,0).string= " en position"
	Sheet.getCellByPosition(3,0).string= "en aspect"
	Sheet.getCellByPosition(4,0).string= "sur planète"
	Sheet.getCellByPosition(5,0).string= "orbe transit"
	Sheet.getCellByPosition(6,0).numberformat=10030 'format date
sheet.charheight=7

'ajustement de la largeur des colonnes feuille transits
	aa=array("99/99/9999","Mercure","99 deg. 99' en Sagittaire","sesqui-carré φ99","Mercure","+ 99 deg. 99'")
		for i=0 to 5
		Sheet.getCellByPosition(i, 1).string=aa(i)
		sheet.columns(i).Optimalwidth = True
		next i
		
'fomat de date = jj/mm/aaaa pour le compteur
	Cell = Sheet.getCellByPosition(6,0)
	cell.numberformat=10030 
	Cell.CellBackColor = RGB(0, 255, 0)

		
'transits progressés ?	
if choix1=2 then goto debut_choix2
	
'************************************************début transits mondiaux************************************************************************************

'ajoute 'mondiaux' au nom
Sheet = Doc.Sheets.getByName(nom_tableau)
	Cell = Sheet.getCellByPosition(7,0)
	if nom_tableau="transits1" then cell.string= cell.string & chr$(10) & "personnelles"
	if nom_tableau="transits2" then cell.string= cell.string & chr$(10) & "mondiaux"

'recalcul des orbes transits en fonction de l'écart réel entre 2 jours dans les éphémérides (lignes 17 et 18)
Sheet = Doc.Sheets.getByName("éphémérides")
	for i=0 to 15
	orbe_transit(i)=abs(Sheet.getCellByPosition(i+1,18).value- Sheet.getCellByPosition(i+1,17).value) '/2
	next i
	

if choix2 ="1" then goto debut1
if choix2 ="2" then goto debut2
	
	
'********pas utilisé (2 fois plus lent): copie du thème (aspects+longitudes) colonnes AE (30) à BJ (61) avec tri croissant des longitudes

'pas utilisé récupération des orbes transits max et de toutes les positions du thème de position(0) à position(239)
	car=0
	'aspects
	for i=1 to 16
	'lecture des orbes feuille éphémérides plutôt que dans "definitions"
	orbe_transit(i-1)=Sheet.getCellByPosition(18, i).getvalue
	if tableau = 1 then orbe_transit(i-1)=orbe_transit(i-1) * 10 'pour les personnelles, orbes multipliés par 10
		'planetes
		for j=1 to 16
		position(car)= Sheet.getCellByPosition(j, i).getvalue
		car=car+1
		next j
	next i
	
'effacement 32 colonnes à partir de AE (sinon erreur à la 2ème lecture)
	for i=0 to 16
		for j=30 to 62 '23 to 26
		sheet.getcellbyposition(j,i).string=""
		next j
	next i
	
'pas utilisé : copie du thème (aspects+longitudes)  colonnes AE-AF,...,BI-BJ
		'longitudes
		for i=1 to 16 '1 to 12
		cell=Sheet.getCellRangeByPosition(i, 0,i, 16)
		aa=cell.getdataArray
		cell=Sheet.getCellRangeByPosition(2*i+29, 0,2*i+29, 16)
		cell.setDataArray(aa)
		next i
		'aspects
		cell=Sheet.getCellRangeByPosition(0, 0,0, 16)
		aa=cell.getdataArray
		for i=30 to 60 step 2
		cell=Sheet.getCellRangeByPosition(i, 0,i, 16)
		cell.setDataArray(aa)
			'en-tête=planète
			sheet.getcellbyposition(i,0).string=sheet.getcellbyposition(i+1,0).string
		next i
			
'pas utilisé : tri par ordre croissant des longitudes pour chaque planète
	'colonnes AE à BI
	for i=30 to 60 step 2
	cell=sheet.getcellrangebyposition(i,1,i+1,16)
	aSortFields(0).Field = 0
	aSortFields(0).SortAscending = true
	'(field = 0 : tri sur 1ère colonne, field = 1 : tri sur 2ème colonne)
	aSortFields(0).Field = 1 
	aSortFields(0).SortAscending = true
	aSortDesc(0).Name = "SortFields"
	aSortDesc(0).Value = aSortFields()
	cell.Sort(aSortDesc())
	next i
	
'pas utilisé : copie du thème dans arrays (longitudes : plan() + aspects : asp())	
for i=1 to 16
	cell=Sheet.getCellRangeByPosition(2*i+29, 1,2*i+29, 16)
	plan(i-1)=cell.getdata
 'arrondi à 1 décimale
'	plan(i-1)=oFA.callFunction("Round",array(plan(i-1),1))
	cell=Sheet.getCellRangeByPosition(2*i+28, 1,2*i+28, 16)
	asp(i-1)=cell.getdataarray
next i



'********utilisé****************



		
		
'****************************début transits mondiaux méthode 1*********************************************************************************************
debut1:
'récupération des longitudes du thème natal
	for i=0 to 15
	natal(i)=Sheet.getCellByPosition(i+1, 1).getvalue
	next i
	
'debut et fin de la plage de lecture éphémérides
'array col_gm() a un maximum de 16368 lignes donc découpage en tronçons de 16000 lignes
	debut=rangminimum : final=debut+16000
	if final >= rangmaximum then final=rangmaximum
'fin de la plage d'écriture dans array col_gm(), début à 0
	final_gm=final-debut
'16000 max
	if final_gm >= 16000 then final_gm=16000

'compteur lignes pour écriture transits dans array tr()
	compte_lignes=0
	
'limites basses et hautes de la barre de progression (si form 'visuel' active)
	if form_ok=1 then
	commande1 = feuille.getControl("ProgressBar1")
	commande1.setrange(0,int(rangmaximum/16000))
	commande2= feuille.getControl("ProgressBar2")
	commande2.setrange(0,final_gm)
	commande2.visible=true
	feuille.getControl("Label3").text=""
	endif
	
'début
'blocs de 16000 lignes max
for k= 0 to int(rangmaximum/16000) 'int(lastrow/16000)
	'affichage barre de progression
	if form_ok=1 then commande1.setvalue(k)
Sheet = Doc.Sheets.getByName("éphémérides")
	'récupération et écriture bloc de données dans arrays col_a(), col_gm() et phase() (16000 lignes)
		'dates
		Cell = Sheet.getCellrangeByPosition(0,debut,0,final)
		col_a()=cell.getDataArray()
		'données planètes
		Cell = Sheet.getCellrangeByPosition(num0,debut,num1,final)
		col_gm()=cell.getData()
				 'arrondi à 1 décimale
				'col_gm=oFA.callFunction("Round",array(col_gm,1))
		'phases
		Cell=Sheet.getCellrangeByPosition(num0+20,debut,num1+20,final)
		phase()=cell.getDataArray()
		
Sheet = Doc.Sheets.getByName(nom_tableau)
		 
  'ligne éphémérides
  for m= 0 to final_gm
  	'affichage barre de progression
	if form_ok=1 then commande2.setvalue(m)		
	 'transitante
	 for i = 0 to num1-num0 
 		 'longitude
		 longitude= col_gm(m)(i)
	  			 			 
			'transitée
			for j= 0 to 15
			gap=longitude-natal(j)
			if gap < 0 then gap=gap+360
		
				'division par 15 pour une approximation de l'aspect
				val1=int(gap/15)
				
				'aspects proches de cette approximation			
				for l=arc(val1,0) to arc(val1,1)
					'orbe à comparer à l'orbe maximum de la transitée
					orbedecimal=gap-angle(l)
				
					  'transit trouvé
					  if abs(orbedecimal) <= orbe_transit(num0 +i-1) then 
		
						'date transit
					  	tr(compte_lignes,0)=col_a(m)(0)
					  	'affichage compteur de date
					  	Sheet.getCellByPosition(6, 0).value =col_a(m)(0)
					  	'transitante
					  	tr(compte_lignes,1)=planete(num0 +i-1)
		 				'transit en degrés minutes + signe
							degres=int(longitude) mod 30
							coeff1=int(longitude/30)*30
							minutes=int(60*(longitude-coeff1-degres))
						abc=str(degres) & chr$(176) & str(minutes) & "' " & signe(coeff1/30) 'ex  12° 23' Lion
						tr(compte_lignes,2)=abc
						'aspect
						tr(compte_lignes,3)=aspect(l mod 16) 'aspect(16)=aspect(0)
						'transitée
						tr(compte_lignes,4)=planete(j)
						'orbe transit en degrés minutes
						signemoins="+"
						if orbedecimal < 0 then orbedecimal=abs(orbedecimal) : signemoins="-"
							degres=int(orbedecimal) mod 30
							coeff1=int(orbedecimal/30)*30
							minutes=int(60*(orbedecimal-coeff1-degres))
						abc=signemoins & str(degres) & chr$(176) & str(minutes) & "'"
						tr(compte_lignes,5)=abc
						'phase
						tr(compte_lignes,6)=phase(m)(i)
					
					compte_lignes= compte_lignes+1
					'saute les aspects restants de la même planète
					exit for
					endif
			'inutile et retarde
			'if gap < angle(k) then exit for
			next l
		 next j
	next i
 next m
	'affichage barre de progression
	if form_ok=1 then commande1.setvalue(k+1)

	'incrémentation des plages de lecture/écriture
	'début
	debut=debut+16000
	if debut > lastrow then exit for
	if debut > rangmaximum then exit for
	'fin
	final=final+16000 : final_gm=16000
	'si fin de fichier, dernier tronçon plus court
	if final > rangmaximum then
		final_gm=16000-(final-rangmaximum)
		final=rangmaximum
	 endif
	 
next k

'écriture dans feuille transits
	redim preserve tr(0 to compte_lignes-1,0 to 6)
	Cell = Sheet.getCellrangeByPosition(0,1,6,compte_lignes)
	cell.setDataArray(tr()) 'setdataarray pour string, nombres et formatage
	
'******************************test

'if compte_lignes > 100 then call filtres(nom_tableau)

'****************************fin test

'mise en couleur
	for i=1 to compte_lignes
		abc= Sheet.getCellByPosition(6,i).getstring
		if abc <> "" then
			car=i+1
			do until car=compte_lignes+2 '+2 est utile si dernière ligne a une couleur
			bcd= Sheet.getCellByPosition(6,car).getstring
				if bcd <> abc or car>compte_lignes then
					cell=Sheet.getCellrangeByPosition(1,i,1,car-1)
					select case abc
						case "phase 1"
						cell.CellBackColor = RGB(167, 236, 106) 'vert
						case "phase 2"
						cell.CellBackColor = RGB(255, 213, 6) 'jaune
						case "phase 3"
						cell.CellBackColor = RGB(255, 153, 153) 'rouge
					end select
					i=car-1
					exit do
				endif
			car=car+1
			loop
		endif
	next i
	
goto final

'********************************************début transits mondiaux méthode 2*****************'***************************************************************
debut2:
	
'thème + aspects
	'planètes
	for i=1 to 16
		'aspects
		for j=1 to 16
		matrice(i,j)=Sheet.getCellByPosition(i,j).getvalue	
		next
	next

Dim oFilterDesc ' Filter descriptor.
Dim oFields(1) As New com.sun.star.sheet.TableFilterField
Dim x As New com.sun.star.table.CellAddress
Dim index1, index2, index3, index4
Dim y(0 to 1000,0) as string 'blocs de données filtrées mises dans feuille temp


		'création feuille éphé2 ou effacement
		If not Doc.Sheets.hasByName("éphé2") Then 
		Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
		Doc.Sheets.insertByName("éphé2", Sheet)
		else
		Sheet = Doc.Sheets.getByName("éphé2")
		sheet.clearcontents (1 or 2 or 4 or 32)
		end if
		'création feuille temp ou effacement
		If not Doc.Sheets.hasByName("temp") Then 
		Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
		Doc.Sheets.insertByName("temp", Sheet)
		else
		Sheet = Doc.Sheets.getByName("temp")
		sheet.clearcontents (1 or 2 or 4 or 32)
		end if
		'récupère index des feuilles éphémérides, éphé2 et temp
		For i = 0 to Doc.Sheets.Count-1
		if Doc.Sheets(i).Name="éphémérides" then index1=i
		if Doc.Sheets(i).Name="éphé2" then index2=i
		if Doc.Sheets(i).Name="temp" then index3=i
		if Doc.Sheets(i).Name=nom_tableau then index4=i
		next i
		'copie partie utile (années choisies) de éphémérides vers éphé2
			'source
			CellRangeAddress.Sheet = index1
			CellRangeAddress.StartColumn = 0
			CellRangeAddress.StartRow = rangminimum
			CellRangeAddress.EndColumn = 12
			CellRangeAddress.EndRow = rangmaximum
			'destination
			x.Sheet = index2
			x.Column = 0
			x.Row = 17
			'copie
			Sheet.copyRange(x, CellRangeAddress)	
			
'début
	'compteur lignes pour écriture feuille transits
	compte_lignes=1
	
	'transitée,Soleil à MC
	for i= 1 to 16
		'affichage transitée feuille transits
		Doc.Sheets.getByName(nom_tableau).getCellByPosition(6,0).string=planete(i-1)
			'aspects
			for j=1 to 16
				'transitante
				for k=num0 to num1
				
Sheet = Doc.Sheets.getByName("éphé2")
		oFilterDesc = Sheet.createFilterDescriptor(True)
		'filtres
			'limite basse
			With oFields(0)
			'.Connection = com.sun.star.sheet.FilterConnection.OR
			.Field = k 'colonne
			.IsNumeric = true
			.NumericValue=matrice(i,j)-orbe_transit(k-1)
			.Operator = com.sun.star.sheet.FilterOperator.GREATER_EQUAL
			end with
			'limite haute
			With oFields(1)
			.Connection = com.sun.star.sheet.FilterConnection.AND
			.Field = k 'colonne
			.IsNumeric = true
			.NumericValue=matrice(i,j)+orbe_transit(k-1)
			.Operator = com.sun.star.sheet.FilterOperator.LESS
			end with
			oFilterDesc.setFilterFields(oFields())
		'options
		oFilterDesc.CopyOutputData = True
		'coodonnées feuille temp pour réception données filtrées
		x.Sheet = index3
		x.Column = 0
		x.Row = 0
		oFilterDesc.OutputPosition = x
		'exécute le filtrage vers temp
		Sheet.filter(oFilterDesc)
		
		'récupère dernière ligne de la feuille temp
Sheet = Doc.Sheets.getByName("temp")
		Curs = Sheet.createCursor
		Curs.gotoEndOfUsedArea(True)
		lastrow = Curs.Rows.Count-1
		'feuille temp vide ? sortie boucle
		if lastrow=0 then goto finboucle
		'2ème et 3ème colonnes : y déplace colonne longitudes de la transitante
			'source
			CellRangeAddress.Sheet = index3
			CellRangeAddress.StartColumn = k
			CellRangeAddress.StartRow = 0
			CellRangeAddress.EndColumn = k
			CellRangeAddress.EndRow = lastrow
			'destination 3ème colonne
			x.Sheet = index3
			x.Column = 2
			x.Row = 0
			'copie
			Sheet.copyRange(x, CellRangeAddress)
			'destination 2ème colonne pour avoir les couleurs
			x.Sheet = index3
			x.Column = 1
			x.Row = 0
			'copie
			Sheet.MoveRange(x, CellRangeAddress)
					'effacement couleurs sauf 2ème colonne (change aussi la taille des caractères er réaffiche les décimales)
					'sheet.getCellrangeByPosition(2,0,4,lastrow).clearcontents (32) 
			'transformation en deg. min. 1er chiffre 3ème colonne
			longitude=sheet.getcellbyposition(2,0).value
			coeff1=int(longitude/30)*30 : coeff2=int(longitude) mod 30
			abc=str(coeff2) & chr$(176) & str(int(60*(longitude-coeff1-coeff2))) & "' " & signe(coeff1/30) 'ex  12° 23' Lion
			'écriture
		redim y(0 to 1000,0)
			for n=0 to lastrow
			y(n,0)=abc
			next n
		redim preserve y(0 to lastrow,0)
			'écriture
			Cell = Sheet.getCellrangeByPosition(2,0,2,lastrow)
			cell.setDataArray(y()) '
		'2ème colonne tout identique : transitante
		redim y(0 to 1000,0)
			for n=0 to lastrow
			y(n,0)=planete(k-1)
			next n
		redim preserve y(0 to lastrow,0)
			'écriture
			Cell = Sheet.getCellrangeByPosition(1,0,1,lastrow)
			cell.setDataArray(y()) '
		'4ème colonne tout identique : aspect
		redim y(0 to 1000,0)
			for n=0 to lastrow
			y(n,0)=aspect(j-1)
			next n
		redim preserve y(0 to lastrow,0)
			'écriture
			Cell = Sheet.getCellrangeByPosition(3,0,3,lastrow)
			cell.setDataArray(y()) 
		'5ème colonne tout identique : transitée
		redim y(0 to 1000,0)
			for n=0 to lastrow
			y(n,0)=planete(i-1)
			next n
		redim preserve y(0 to lastrow,0)
			'écriture
			Cell = Sheet.getCellrangeByPosition(4,0,4,lastrow)
			cell.setDataArray(y()) 
		'copie temp dans feuille transits	
			'source
			CellRangeAddress.Sheet = index3
			CellRangeAddress.StartColumn = 0
			CellRangeAddress.StartRow = 0
			CellRangeAddress.EndColumn = 4
			CellRangeAddress.EndRow = lastrow
			'destination
			x.Sheet = index4
			x.Column = 0
			x.Row = compte_lignes
			'copie
			Sheet.copyRange(x, CellRangeAddress)
				
		compte_lignes=compte_lignes+lastrow+1
			
finboucle:
			next k	
		next j
	next i		


'tri par ordre croissant colonne date
Sheet = Doc.Sheets.getByName(nom_tableau)
	cell=sheet.getcellrangebyposition(0,1,4,compte_lignes+1)
	aSortFields(0).Field = 0
	aSortFields(0).SortAscending = true
	'(field = 0 : tri sur 1ère colonne, field = 1 : tri sur 2ème colonne)
	aSortDesc(0).Name = "SortFields"
	aSortDesc(0).Value = aSortFields()
	cell.Sort(aSortDesc())

goto final		
'***********************************************fin transits mondiaux méthode 2*******************************		



'***********************************************fin transits mondiaux**********************************************************************************************

erreur:
msgbox "compteur : " & str(m) & " - i : " & str(i) & " - compte_lignes : " & str(compte_lignes)
exit sub

'**********************************************début tranits progressés********************************************************************************
debut_choix2:

dim annee_naissance as string
dim date_num as long
dim date_transit as string
dim ecart_jour as double, ecart_jour2 as double, ecart_global as double
dim heure_naissance as double
dim janvier_naissance as long
dim janvier_annee as long
dim jour_naissance as string
dim ligne as long
dim longitude2 as double
dim mois_naissance as integer
dim orbe_jour as double
dim transitante as string
dim ecart(1 to 40) as double 'ecart annuel
dim ecart_h(1to 40) as double
dim ecart_j(1to 40) as double
dim long_signe(11) as double
dim longitudes(1 to 40) as double 'transitantes et transitées
dim natales(1 to 40) as double 'transitées
dim maison(1 to 12) as double


	'limites basses et hautes de la barre de progression (si form 'visuel' active)
	if form_ok=1 then
	commande1 = feuille.getControl("ProgressBar1")
	commande1.setrange(val(anneemin),val(anneemax))
	feuille.getControl("ProgressBar2").visible=false
	feuille.getControl("Label3").text=""
	endif
	
	
Sheet = Doc.Sheets.getByName("éphémérides")
'date de naissance
	abc=Sheet.getCellByPosition(22, 1).getstring		
	date_naissance=datevalue(abc)
	if date_naissance<0 then msgbox "impossible, date de naissance < 1900" : exit sub
	annee_naissance=year(abc)
'	if annee <= val(annee_naissance) then msgbox "erreur, choisir année supérieure à " & annee_naissance : exit sub
	mois_naissance=month(abc)	
	jour_naissance=day(abc)
	
'heure de naissance
	abc=Sheet.getCellByPosition(22, 2).getstring
	 'timevalue=coefficient de 0 à 1 (midi=0,5, minuit=1)
	heure_naissance=timevalue(abc)

Sheet = Doc.Sheets.getByName(nom_tableau)
	'écriture nom
	Cell = Sheet.getCellByPosition(7,0)
		cell.string= cell.string  & chr$(13) & "progressé_" & chr$(13)
		if choix_pro_natal=1 then
		cell.string=cell.string & "natal"
		else
		cell.string=cell.string & "progressé"
		endif
	Cell.CellBackColor = RGB(0, 255, 0)
	'fomat de date = jj/mm/aaaa pour le compteur
	Cell = Sheet.getCellByPosition(6,0)
	cell.numberformat=10030 
	Cell.CellBackColor = RGB(0, 255, 0)
	
'ajustement de la hauteur 1ère ligne
	Sheet.Rows(0).OptimalHeight = True

'cuspides des Maisons natales
Sheet = Doc.Sheets.getByName("éphémérides")	
	for i=1 to 12
	maison(i)=Sheet.getCellByPosition(21,i).getvalue
	next i	
	if maison(1)=0 then message="pas de maisons dans la feuille éphémérides!"
	
'estimation de l'augmentation annuelle des longitudes des cuspides maisons pour le progressé/progressé
	aa=array(0.72,0.79,0.895,0.955,0.91,0.81,0.72,0.79,0.895,0.955,0.91,0.81)
	
'longitudes des signes
	for i=0 to 11
	long_signe(i)=i*30
	next i

'si progressé/natal, regroupement des données natales dans natales() : (planètes 1-16, cuspides 17-28 et signes 29-44)
	if choix_pro_natal=1 then
		for i=1 to 40
			'planètes 1-16
			if i <= 16 then natales(i)=Sheet.getCellByPosition(i, 1).getvalue
			'cuspides 17-28
			if i > 16 and i <= 28 then natales(i)=maison(i-16)
			'signes 29-44
			if i > 28 then natales(i)=long_signe(i-29)
		next i
	endif
	
'remise à 0 des longitudes
	for i=1 to 40
	longitudes(i)=0
	next i

'valeur numérique du 1er janvier précédant l'anniversaire
	janvier_naissance=datevalue("01/01/" & annee_naissance)
		
'****debut*****
	orbe_transit(0)=0.02
	'compteur ligne feuille transit	
	compte_lignes=1
	
	on error goto erreur2
'anneés		
for l=val(anneemin) to val(anneemax)
	'barre de progression si form 'visuel' active
	if form_ok=1 then commande1.setvalue(l)
		
	'valeur numérique du 1er janvier de l'année analysée
	janvier_annee=datevalue("01/01/" & l)

	Sheet = Doc.Sheets.getByName("éphémérides")
	'détermine la ligne éphémérides correspondant à l'année de calcul du thème progressé (1 jour= 1 an)
	ligne=date_naissance -datevalue(Sheet.getCellByPosition(0, 17).getstring) + 17 + l -val(annee_naissance)
	if ligne < 17 or ligne > lastrow then msgbox str(l) & " : " & "année hors plage éphémérides" : exit sub
	
	'offsets >0 ou <0 à appliquer aux longitudes des planètes transitantes et des planètes et maisons transitées (en progressé/progressé))
	'****ecart_heure (ecart_h()) est très important sinon possibilité de décalage de plusieurs mois si naissance l'après-midi !!
	for i=1 to 28
		if i <= 12 then
			'offset annuel planètes (i=1-12) : écart entre 2 lignes (en progressé = 1 an, /12 = 1 mois, /365 = 1 jour)
			ecart(i)=Sheet.getCellByPosition(i, ligne+1).getvalue-Sheet.getCellByPosition(i, ligne).getvalue
			'compensation si passage à 360 entre 2 lignes
				'planète rétrograde
				if ecart(i) > 300 then ecart(i)=ecart(i)-360
				'planète directe
				if ecart(i) < -300 then ecart(i)=ecart(i)+360
			'ecart_heure, offset pour tenir compte de l'heure de naissance
			ecart_h(i)=heure_naissance*ecart(i)
		elseif i > 16 then
			'offset annuel maisons (i=17-28)
			ecart(i)=aa(i-17)
		endif
		'ecart_janvier : offset pour tenir compte du nombre de jours entre le  1er janvier et le jour de naissance (ex 30/10 : 300)
		ecart_j(i)=(date_naissance-janvier_naissance)*ecart(i)/365
	next i


	'récupération des longitudes au 1er janvier pour 1 ligne = 1 année (planètes 1-16, cuspides 17-28 et signes 29-44)
		for j=40 to 1 step -1 '-1 permet de mettre les longitudes des axes (j=13-16) après celles des maisons (j=17-28) en progressé
			
			'récupération des données progressées
					'planètes 13-16 =AS,FC,DS,MC : même valeurs que les maisons correspondantes (1,4,7,10 soit j=17,20,23,26)
				if j=16  then
					longitudes(13)=longitudes(17) 'ASC
					longitudes(14)=longitudes(20) 'FC
					longitudes(15)=longitudes(23) 'DS
					longitudes(16)=longitudes(26) 'MC
				endif
				'planètes 1-12
				if j <= 12 then
					'longitude au 1er janvier
				 	longitudes(j)=Sheet.getCellByPosition(j,ligne).getvalue + ecart_h(j) - ecart_j(j) 
				 		if longitudes(j) >= 360 then longitudes(j)=longitudes(j)-360
						if longitudes(j) < 0 then longitudes(j)=longitudes(j)+360
				'maisons j=17-28 (approximatif à 2 mois près parfois !)
				elseif j > 16 and j <= 28 then
					'offset global depuis l'année de naissance / maison natale
					ecart_global=ecart(j)*(l -val(annee_naissance))
					'longitude au 1er janvier
					longitudes(j)=maison(j-16) + ecart_global - ecart_j(j) 
						if longitudes(j) > 360 then longitudes(j)=longitudes(j)-360
						if longitudes(j) < 0 then longitudes(j)=longitudes(j)+360
				'signes j=29-44 pas d'offset
				elseif j > 28 then
					longitudes(j)=long_signe(j-29)
				endif
		next j
	
 'boucles de comparaison transitantes/transitées pour chaque jour de l'année
	'transitantes Soleil à Saturne
	for i=1 to 7
		'orbe
		orbe_jour=abs(ecart(i)/365)
		'offset journalier (positif ou négatif si planète rétrograde) de la transitante
		ecart_jour=ecart(i)/365
	
	
		'transitées natales ou progressées : planètes 1-16, cuspides 17-28 et signes 29-44
		for j=1 to 40
			'pas de comparaison d'une planète à elle-même
			if j=i then goto fin_j
			'progressé/natal
			if choix_pro_natal=1 then
				'pas d'offset journalier
				ecart_jour2=0
				'longitude transitée au 1er janvier
				longitude2= natales(j)
			'progressé/progressé
			else
				'offset journalier, nul(signes), positif ou négatif si planète rétrograde(maisons, planètes)
				ecart_jour2=ecart(j)/365
				'longitude transitée au 1er janvier
				longitude2= longitudes(j)
			endif
		
		
	'valeur numérique du 1er janvier
	date_num=janvier_annee
	'longitude transitante au 1er janvier '31 décembre de l'année précédente (augmente d'un jour les dates de transits)
	'***ne pas déplacer au-dessus de for j=... !!!
	longitude= longitudes(i)'- ecart_jour
	
			'jours de l'année
			for k=1 to 365
			
			'transitantes Soleil à Saturne
			longitude= longitude + ecart_jour
				if longitude >= 360 then longitude=longitude-360
				if longitude < 0 then longitude=longitude+360
							 	 				
			'transitées Soleil à MC, cuspides, signes
			longitude2= longitude2 + ecart_jour2
				if longitude2 >= 360 then longitude2=longitude2-360
				if longitude2 < 0 then longitude2=longitude2+360
								
			'mesure différence transitante-transitée
			gap=longitude - longitude2							
				if gap < 0 then gap=gap+360
			'division par 15 pour une approximation de l'aspect
			val1=int(gap/15)
			'on ne garde que les conjonctions (val1=0) pour les transits sur cuspides maisons (j de 17 à 28) ou signes (j de 29 à 40)
				if j >16  then if val1 >0 then goto fin_k
						
				'aspects proches de cette approximation			
				for m=arc(val1,0) to arc(val1,1)
			  	orbedecimal=gap-angle(m)
					  
				  '*********transit trouvé**********
				  if abs(orbedecimal) <= orbe_jour then
					  	Sheet = Doc.Sheets.getByName(nom_tableau)
						transitante=planete(i-1)
						'position transit en degrés minutes + signe
						coeff1=int(longitude/30)*30 : coeff2=int(longitude) mod 30
						abc=str(coeff2) & chr$(176) & str(int(60*(longitude-coeff1-coeff2))) & "' " & signe(coeff1/30) 'ex  12° 23' Lion
						'orbe transit en degrés minutes
						signemoins="+"
						if orbedecimal < 0 then orbedecimal=abs(orbedecimal) : signemoins="-"
						coeff1=int(orbedecimal/30)*30 : coeff2=int(orbedecimal) mod 30
						bcd=signemoins & str(coeff2) & chr$(176) & str(int(60*(orbedecimal-coeff1-coeff2))) & "'"
								
						'écriture ligne dans tableau progressé
							'compteur de date (convertie du format numérique au format chaïne)
							Sheet.getCellByPosition(6, 0).string = cdate(date_num)
							'date au format numérique (important pour la suite : tableau par planete !)
							Sheet.getCellByPosition(0, compte_lignes).value = date_num
							'transitante
							Sheet.getCellByPosition(1, compte_lignes).string= transitante
							'position degrés + signe
							Sheet.getCellByPosition(2, compte_lignes).string= abc
							'aspect
							Sheet.getCellByPosition(3, compte_lignes).string= aspect(m mod 16)
							'transitée
								'planète (1 à 16)
								if j <= 16 then Sheet.getCellByPosition(4, compte_lignes).string= planete(j-1)
								'cuspide (17 à 28) et signe (29 à 40)
								if j >16  then Sheet.getCellByPosition(4, compte_lignes).string=str(j-16)
							'orbe
							Sheet.getCellByPosition(5, compte_lignes).string= bcd
																		
						compte_lignes= compte_lignes+1
						Sheet = Doc.Sheets.getByName("éphémérides")
						goto fin_j
					endif
		 		 next m
fin_k:
			 date_num=date_num+1
			next k
fin_j:
		next j	
	next i
next l


'***********************************************fin transits progressés*************************************************************************

final:

'centrage caractères
	Sheet = Doc.Sheets.getByName(nom_tableau)
	sheet.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	
'format date pour la 1ère colonne
	Sheet.getCellrangeByPosition(0,1,0,compte_lignes+1).numberformat=10030
	
'écriture plage de dates
	cell=Sheet.getCellByPosition(8, 0)
	abc=year(Sheet.getCellByPosition(0, 1).getstring)
	bcd=year(Sheet.getCellByPosition(0,compte_lignes-1).getstring)
	cell.string=abc
	if bcd<>abc then cell.string=cell.string & "-" & bcd
	Cell.CellBackColor = RGB(0, 255, 0)
	
'largeur colonnes nom et dates
	sheet.columns(7).Optimalwidth = True
	sheet.columns(8).Optimalwidth = True

'affichage nom et dates dans la form (si form 'visuel' active)
	if form_ok=1 then
		if nom_tableau="transits2" then abc="Label4"
		if nom_tableau="progressé" then abc="Label5"

		if nom_tableau="transits1" then abc="Label9"
	commande1 = feuille.getControl(abc)
	commande1.text=Sheet.getCellByPosition(7,0).getstring & "  " & Sheet.getCellByPosition(8,0).getstring
	endif
			
'mise au 1er plan de la feuille
	Sheet = Doc.Sheets.getByName(nom_tableau)
	Controller = Doc.CurrentController
	controller.setActiveSheet(sheet)
	
'active boutons tableaux, thème, aspects, etc.
call actions_boutons2(true)	

msgbox "terminé : " & compte_lignes & " lignes écrites"
exit sub

erreur2:
	msgbox "erreur ! " & " gap : " & str(gap) & "  val1 : " & str(val1)

End sub



Sub filtres(nom_tableau)
Dim oFilterDesc ' Filter descriptor.
Dim oFields(2) As New com.sun.star.sheet.TableFilterField
Dim x As New com.sun.star.table.CellAddress
Dim index1, index2, index3, index4
Dim aa
Dim teinte

	'création feuille temp ou effacement
		If not Doc.Sheets.hasByName("temp") Then 
		Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
		Doc.Sheets.insertByName("temp", Sheet)
		else
		Sheet = Doc.Sheets.getByName("temp")
		sheet.clearcontents (1 or 2 or 4 or 32)
		end if
	'création feuille transits ou effacement
		If not Doc.Sheets.hasByName("transits") Then 
		Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
		Doc.Sheets.insertByName("transits", Sheet)
		else
		Sheet = Doc.Sheets.getByName("transits")
		sheet.clearcontents (1 or 2 or 4 or 32)
		end if
	'récupère index des feuilles transits2, temp, transits
		For i = 0 to Doc.Sheets.Count-1
		if Doc.Sheets(i).Name=nom_tableau then index1=i
		if Doc.Sheets(i).Name="temp" then index2=i
		if Doc.Sheets(i).Name="transits" then index3=i
		next i
		
		aa=array("phase 1","phase 2","phase 3")
		teinte=array(RGB(167, 236, 106),RGB(255, 213, 6),RGB(255, 153, 153))
		compte_lignes=1
		
		for i=0 to 3
Sheet = Doc.Sheets.getByName(nom_tableau)
		oFilterDesc = Sheet.createFilterDescriptor(True)
		Sheet.filter(oFilterDesc)
		'filtres
		oFilterDesc = Sheet.createFilterDescriptor(True)
			if i < 3 then
				With oFields(0)
				'.Connection = com.sun.star.sheet.FilterConnection.OR
				.Field = 6 'colonne
				.IsNumeric = false
				.StringValue=aa(i)
				.Operator = com.sun.star.sheet.FilterOperator.EQUAL
				end with
			else
			Sheet = Doc.Sheets.getByName("temp")
			sheet.clearcontents (1 or 2 or 4 or 32)
			Sheet = Doc.Sheets.getByName(nom_tableau)
				for j=0 to 2
					With oFields(j)
					'.Connection = com.sun.star.sheet.FilterConnection.OR
					.Field = 6 'colonne
					.IsNumeric = false
					.StringValue=aa(j)
					.Operator = com.sun.star.sheet.FilterOperator.NOT_EQUAL
					end with
				next j
			endif
		oFilterDesc.setFilterFields(oFields())
		'options
		oFilterDesc.ContainsHeader = True
		oFilterDesc.CopyOutputData = True 'plante si 3 filtres à NOT_EQUAL
		'coodonnées feuille temp pour réception données filtrées
		x.Sheet = index2
		x.Column = 0
		x.Row = 0
		oFilterDesc.OutputPosition = x
		'exécute le filtrage vers temp
		Sheet.filter(oFilterDesc)
	goto fin
		'dernière ligne de temp
Sheet = Doc.Sheets.getByName("temp")
		Curs = Sheet.createCursor
		Curs.gotoEndOfUsedArea(True)
		lastrow = Curs.Rows.Count-1
		'mise en couleur
		if i < 3 then Sheet.getCellRangeByPosition(1,0,1,lastrow).CellBackColor = teinte(i)
		'copie temp dans feuille transits	
			'source = temp
			CellRangeAddress.Sheet = index2
			CellRangeAddress.StartColumn = 0
			CellRangeAddress.StartRow = 0
			CellRangeAddress.EndColumn = 6
			CellRangeAddress.EndRow = lastrow
			'destination = transits
			x.Sheet = index3
			x.Column = 0
			x.Row = compte_lignes
			'copie
			Sheet.copyRange(x, CellRangeAddress)
			
	compte_lignes=compte_lignes+lastrow+1
fin:
	next i
End Sub


Sub ecrit_couleurs(nom_tableau,filtre())
Dim oFilterDesc ' Filter descriptor.
Dim oFields(0) As New com.sun.star.sheet.TableFilterField
Dim x As New com.sun.star.table.CellAddress


Sheet = Doc.Sheets.getByName("temp")
	sheet.clearcontents (1 or 2 or 4 or 32)
	lastrow=0
	
	For i=0 to 2
Sheet = Doc.Sheets.getByName(nom_tableau)
		oFilterDesc = Sheet.createFilterDescriptor(True)
		'filtres
		With oFields(0)
		.Field = 6 'colonne
		'.IsNumeric = true'false
		'.NumericValue=0
		'.StringValue = 
		.Operator = com.sun.star.sheet.FilterOperator.EMPTY
		end with
		oFilterDesc.setFilterFields(oFields())
		'options
		oFilterDesc.CopyOutputData = True
'Sheet = Doc.Sheets.getByName("temp")
	'	cell=Sheet.getCellByPosition(0,lastrow)
	x.Sheet = 3
	x.Column = 0
	x.Row = 0
		oFilterDesc.OutputPosition = x 'Doc.Sheets.getByName("temp").getCellByPosition(0,lastrow)
		'oFilterDesc.ContainsHeader = True
	'exécute
Sheet = Doc.Sheets.getByName(nom_tableau)
		Sheet.filter(oFilterDesc)
		
		'récupère le nombre de lignes utilisées dans la feuille éphémérides
Sheet = Doc.Sheets.getByName("temp")
		Curs = Sheet.createCursor
		Curs.gotoEndOfUsedArea(True)
		lastrow = Curs.Rows.Count
		'couleur
		Sheet.getCellRangeByPosition(1,0,1,lastrow-1).CellBackColor = RGB(167, 236, 106) 'vert
	Next i
		
		
'effacement filtre
'oFilterDesc = Sheet.createFilterDescriptor(True)
'Sheet.filter(oFilterDesc)
End Sub




sub tableau_par_planete

Dim Col As Object
dim doc2 as object
Dim Row As Object
Dim aBorder as New com.sun.star.table.BorderLine
Dim oBorder as New com.sun.star.table.TableBorder
Dim filterArgs(1) as new com.sun.star.beans.PropertyValue

dim a1%,a2%
dim anneemax as string
dim anneemin as string
dim aspect_transit as string
dim chaine as string
dim colonne as integer
dim couleur_planete as long
dim date_transit as string
dim decalage as integer
dim indice%
dim ligne as integer
dim mini as long, maxi as long
dim nom_tableau as string
dim nom_theme as string
dim phase_transit as string
dim pos1%
dim rangmin as long, rangmax as long
dim refanneemax as string
dim repere as integer
dim transitante as string
dim transitee as string

dim an(0 to 1000) as string
dim annum(0 to 1000) as long
dim opt(2) as string
dim tr(2) as string
dim yy(0,0 to 15) 'permet d'écrire une ligne dans une feuille ex range(0,0,15,0)
dim zz() as integer

'on error resume next
Doc = ThisComponent

'vérifie la présence de feuilles de transits
tr=array("transits2","transits1","progressé")
opt=array("- les mondiales (1) "," - les personnelles (2) "," - le thème progressé (3)")
abc="tableaux par planètes pour "
car=0

	for i= 0 to 2
	If Doc.Sheets.hasByName(tr(i)) Then 
		abc=abc & opt(i)
		if car = 0 then car=i+1
		endif
	next i
if car=0 then msgbox "pas de feuilles de transits, exécuter : tableaux des transits" : exit sub

'choix transits
ref0:
car=val(inputbox (abc , "tableau transits par planètes", car))
if car=0 then exit sub
if car < 1 or car > 3 then goto ref0
	if car=2 then
	num0=2
	num1=4
	nom_tableau="transits1"
	abc="personnelles"
	elseif car=1 then
	num0=5
	num1=11
	nom_tableau="transits2"
	abc="mondiales"
	else
	num0=0
	num1=6
	nom_tableau="progressé"
	abc="thème progressé"
	endif
	
'vérifie si feuille transits présente
	If not Doc.Sheets.hasByName(nom_tableau) Then msgbox "pas de feuille transits nommée : " & nom_tableau : exit sub
	Sheet = Doc.Sheets.getByName(nom_tableau)
	if Sheet.getCellByPosition(0, 1).getstring ="" then  msgbox "le tableau des transits est vide, exécuter 'tableau_transits' d'abord" : exit sub

'récupère le nombre de lignes utilisées dans la feuille transits
	Curs = Sheet.createCursor
	Curs.gotoEndOfUsedArea(True)
	lastrow = Curs.Rows.Count -1

'détermination années min et max du tableau des transits
	anneemin = right (Sheet.getCellByPosition(0, 1).getstring,4) 
	anneemax = right (Sheet.getCellByPosition(0, lastrow).getstring,4) 
	refanneemax=anneemax
	if anneemin=anneemax then goto ref3

'choix des années de début et fin
ref1:
	choix=inputbox ("entrer année de départ" , "tableau transits " & abc & anneemin & " - " & anneemax, anneemin)
	if choix="" then exit sub
	if choix < anneemin or choix > anneemax then goto ref1
anneemin = choix
ref2:
	choix=inputbox ("entrer année de fin" , "tableau transits " & abc & anneemin & " - " & anneemax, anneemax)
	if choix="" then exit sub
	if choix < anneemin or choix > anneemax then goto ref2
anneemax = choix

'confirmation
ref3:
	choix=inputbox (" création des tableaux de transits " & abc & " pour les années " & anneemin & " - " & anneemax &" (O,N) ?", "confirmation","O")
	if choix <> "O" then exit sub

'désactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(false)

'définitions variables aspect et planete du theme
call definitions

'suppression si existent des feuilles planètes + feuille maisons + feuille signes
for i=num0 to num1 ' soleil à pluton ou jupiter à lilith
	abc=planete(i)
	If Doc.Sheets.hasByName(abc) Then Doc.Sheets.RemoveByName(abc)'suppression feuille
next i
	If Doc.Sheets.hasByName("maisons") Then Doc.Sheets.RemoveByName("maisons")
	If Doc.Sheets.hasByName("signes") Then Doc.Sheets.RemoveByName("signes")
	
'création 1ère feuille planète
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
 	Doc.Sheets.insertByName(planete(num0), Sheet)

'création feuilles maisons et signes
	tr=array("maisons","signes")
	for i=0 to 1
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
 	Doc.Sheets.insertByName(tr(i), Sheet)
	'lignes bleues autour des cellules
	call cadre_cellules(tr(i),7,12) 'feuille, dernière colonne, dernière ligne
	next i
	
'récupération nom du thème + dates min et max dans feuille transits
Sheet = Doc.Sheets.getByName(nom_tableau)
	bcd=Sheet.getCellByPosition(7,0).getstring
	'écrit la plage de dates des transits
	nom_theme= bcd & chr$(13) & anneemin
	if anneemax <> anneemin then nom_theme=nom_theme & " - " & anneemax
		
'configuration 1ère feuille planète (en-têtes, couleurs)
	abc=planete(num0)
Sheet = Doc.Sheets.getByName(abc)
	sheet.charheight=6

	'alignement du texte : centre H et V
	cell=Sheet.getCellRangeByPosition(0,0,17,79) 
	Cell.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	Cell.VertJustify = com.sun.star.table.CellVertJustify.CENTER
				
	'lignes bleues autour des cellules
	call cadre_cellules(abc,17,80) 'feuille, dernière colonne, dernière ligne
	
	'écriture en array des transitées yy(0,ligne)
	for j=0 to 15
	yy(0,j)=planete(j)
	next j
	
	'en-têtes colonnes et lignes
	for i =0 to 15
		'1ère ligne : en-têtes colonnes des planètes transitées
		Cell = Sheet.getCellRangeByPosition(1,5*i,16,5*i)
		cell.setdataarray(yy)
		cell.CellBackColor = RGB(0,204,204)
		'ligne 1 transitante
		Cell = Sheet.getCellByPosition(0,5*i) : cell.string=abc 'colonne de gauche
		Cell.CellBackColor = RGB(153,204,255) 'bleu
		Cell = Sheet.getCellByPosition(17,5*i) : cell.string=abc 'colonne de droite
		Cell.CellBackColor = RGB(153,204,255) 'bleu
		'ligne 2 aspect
		Cell = Sheet.getCellByPosition(0,5*i+1) : cell.string=aspect(i) 'colonne de gauche
		Cell.CellBackColor = RGB(153,204,255) 'bleu
		Cell = Sheet.getCellByPosition(17,5*i+1) : cell.string=aspect(i) 'colonne de droite
		Cell.CellBackColor = RGB(153,204,255) 'bleu
		'ligne 3 phase 1
		Sheet.getCellByPosition(0,5*i+2).string="ph.1" 'colonne de gauche
		Sheet.getCellByPosition(17,5*i+2).string="ph.1" 'colonne de droite
		Sheet.getCellRangeByPosition(0,5*i+2,17,5*i+2).CellBackColor = RGB(167, 236, 106) 'vert
		'ligne 4 phase 2
		Sheet.getCellByPosition(0,5*i+3).string="ph.2R" 'colonne de gauche
		Sheet.getCellByPosition(17,5*i+3).string="ph.2R" 'colonne de droite
		Sheet.getCellrangeByPosition(0,5*i+3,17,5*i+3).CellBackColor = RGB(255, 213, 6) 'orange
		'ligne 5 phase 3
		Sheet.getCellByPosition(0,5*i+4).string="ph.3" 'colonne de gauche
		Sheet.getCellByPosition(17,5*i+4).string="ph.3" 'colonne de droite
		Sheet.getCellrangeByPosition(0,5*i+4,17,5*i+4).CellBackColor = RGB(255, 153, 153) 'rouge
	next i	
			
	'1ère cellule : nom + dates
	cell=Sheet.getCellByPosition(0,0)
	cell.string=nom_theme
	Cell.CellBackColor = RGB(255,255,255)
	Cell.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER

'création des autres feuilles planètes à partir de la 1ère
	for k=num0+1 to num1
	doc.sheets.copybyname(planete(num0),planete(k),thisComponent.getSheets.getCount)
	next k

	'modification du nom de la planète dans chaque feuille
	for k=num0+1 to num1
	abc=planete(k)
Sheet = Doc.Sheets.getByName(abc)
		'1ère ligne, dernière colonne
		Sheet.getCellByPosition(17,0).string=abc
		'lignes suivantes de 5 en 5
	  	for i =1 to 15
		Sheet.getCellByPosition(0,5*i).string=abc
		Sheet.getCellByPosition(17,5*i).string=abc
	  	next i
	next k

'configuration feuilles maisons et signes
tr=array("maisons","signes")
	for j=0 to 1
		Sheet = Doc.Sheets.getByName(tr(j))
		sheet.charheight=6
		Sheet.getCellByPosition(0,0).string=nom_theme
		'1ère ligne en bleu
		Sheet.getCellRangeByPosition(1,0,7,0).CellBackColor =  RGB(153,204,255) 'bleu clair
		'1ère colonne en bleu
		Sheet.getCellRangeByPosition(0,1,0,12).CellBackColor =  RGB(153,204,255) 'bleu clair
		'1ère ligne en-têtes planètes
		for i=1 to 7
			Sheet.getCellByPosition(i,0).string=planete(i-1)
		next i
		'1ère colonne en-têtes
		for i= 1 to 12
			'maisons
			if j=0 then Sheet.getCellByPosition(0,i).string ="Maison " & i
			'signes
			if j=1 then Sheet.getCellByPosition(0,i).string =signe(i-1)
		next i
		cell=Sheet.getCellRangeByPosition(0,0,7,12) 
		Cell.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
		Cell.VertJustify = com.sun.star.table.CellVertJustify.CENTER
	next j

'détermination des lignes min et max dans la feuille transit
Sheet = Doc.Sheets.getByName(nom_tableau)
	cell=Sheet.getCellRangeByPosition(0, 1,0, lastrow)
	'récupération des dates dans array zz
	zz()=cell.getdataarray()
	'dates mini et maxi à rechercher
	mini=datevalue("01/01/" & anneemin)
	maxi=datevalue("31/12/" & anneemax)
	
	'détermination de la première ligne 
		for i=0 to ubound(zz)
		car=zz(i)(0)
		if car >= mini then rangmin=i+1: exit for
		next i
	'détermination de la dernière ligne
	if anneemax=refanneemax then rangmax=lastrow : goto debut
		for j=i to ubound(zz)
		car=zz(j)(0)
		if car > maxi then rangmax=j: exit for
		next j

debut:
'limites basses et hautes de la barre de progression (si form 'visuel' active)
	if form_ok=1 then
	commande1 = feuille.getControl("ProgressBar1")
	commande1.setrange(rangmin,rangmax)
	commande2= feuille.getControl("ProgressBar2")
'	commande2.setrange(0,final_gm)
	commande2.visible=false
	feuille.getControl("Label3").text=""
	endif


'************début****************
for j=rangmin to rangmax
	'affichage barre de progression (si form 'visuel' active)
	if form_ok=1 then commande1.setvalue(j)
Sheet = Doc.Sheets.getByName(nom_tableau)
	'lecture date 1ere colonne
	date_transit=Sheet.getCellByPosition(0, j).getstring
	'compteur
	Sheet.getCellByPosition(6, 0).string = date_transit			
	'lecture planète transitante
	cell=Sheet.getCellByPosition(1, j)
	transitante = cell.getstring
	couleur_planete=Cell.CellBackColor
	'lecture planète transitée
	transitee=Sheet.getCellByPosition(4,j).getstring
	
	'transitée = maison ou signe ? (uniquement en progressé, val(transitee) étant un nombre il s'agit forcément d'une maison(1 à 12) ou d'un signe(13 à 24)
	if val(transitee) >= 1 and val(transitee) <= 24 then
		'dans ce cas écriture dans la feuille maisons, pas dans les feuilles planètes
		'cuspides = transitees de 1 à 12
		if val(transitee) <=12 then k=0 : Sheet = Doc.Sheets.getByName("maisons")
		'signes = transitees de 13 à 24 donc k=12 chiffre à soustraire pour écrre dans les lignes 1 à 12
		if val(transitee) >12 then k=12 : Sheet = Doc.Sheets.getByName("signes")
			'écriture date transit
			for i=1 to 7
				if Sheet.getCellByPosition(i, 0).string=transitante then
					cell=Sheet.getCellByPosition(i, val(transitee)-k)
					abc=cell.string : if abc<>"" then abc=abc & chr(13)
				 	cell.string=abc & date_transit
				 	exit for
			 	endif
			next i
		goto fin_j
		endif
		
	'lecture aspect
	aspect_transit=Sheet.getCellByPosition(3, j).getstring

	'decalage de 1 à 3 lignes à prévoir dans la feuille à écrire, suivant la couleur
	decalage=0
	select case couleur_planete
	case RGB(167, 236, 106) 'vert
	decalage=1
	case RGB(255, 213, 6)  'orange
	decalage=2
	case RGB(255, 153, 153) 'rouge
	decalage=3		 
	end select
	
	'détermine la colonne planète appropriée dans la feuille à écrire
	for i=0 to 15
	if transitee=planete(i) then colonne=i+1 : exit for
	next i	
	'détermine la ligne aspect appropriée dans la feuille à écrire
	for i=0 to 15
	if aspect_transit=aspect(i) then ligne=5*i+1+decalage : exit for
	next i
	if colonne=0 or ligne=0 then msgbox "erreur avec transit de " & transitante & " " & aspect_transit & " " & transitee & " le " & date_transit & chr$(13) & "ligne " & ligne & " colonne " & colonne
	'écriture date transit
Sheet = Doc.Sheets.getByName(transitante)
	Cell = Sheet.getCellByPosition(colonne, ligne)
	if cell.string = "" then
	cell.string=date_transit
	else
	cell.string=cell.string & chr$(13) & date_transit
	endif
fin_j:
next j


'compression par élimination des dates contigues
	'mise à 0 arrays
	for i=0 to 1000
	an(i)=""
	annum(i)=0
	next i
	
	'transitantes
	for j=num0 to num1
Sheet = Doc.Sheets.getByName(planete(j))
	 for colonne=1 to 16
	  for ligne=1 to 80
		abc=Sheet.getCellByPosition(colonne, ligne).getstring
		if abc="" then goto fin1
		pos1=0
		if len(abc) > 10 then 'plusieurs dates dans la cellule, séparées par chr$(13)
			'récupératon dates sous forme complète (an(i)) et numérique (annum(i))
			car1=int(len(abc)/11)
				for i=0 to car1
				bcd=mid$(abc,11*i+1,10)
				'date complète
				an(i)=bcd
				if bcd="" then exit for
				'date au format numérique (il y avait un bug avec la date 02/04/1945 théoriquement datevalue=16529, qui génèrait une erreur)
				annum(i)=datevalue(bcd) 
				next i
			'chaine=une ou plusieurs dates moyennes de plusieurs dates contigues
			repere=0 : chaine=""
				for i=0 to car1
					if annum(i+1)-annum(i)=1 then '2 dates contigues
	 				pos1=pos1+1
	 				else
	 				'ne pas mettre cette boucle sur une ligne sinon pos1 ensuite n'est pas remis à 0
	 				'ne pas regrouper a1,a2,indice sinon indice donne -1 au  lieu de 0 !
	 					'pos1/2 est arrondi au-dessus ex. 1/2=1, 3/2=2
	 					a1=pos1/2
	 					'pos1 mod 2 pour reculer d'un cran si pos1 est impair cad si nombre de dates est pair (ex pour 4 dates, on prend la 2ème pas la 3ème)
	 					a2=pos1 mod 2
	 					indice=repere+a1-a2
						if chaine<>"" then
						chaine= chaine+chr$(13)+an(indice)
						else
						chaine=an(indice) 'enlève le CR de début
						endif
					pos1=0 
					repere=i+1 'ne pas mettre sur une ligne avec la précédente sinon répétition de dates !
					endif
				next i
			'cas particulier ?	
			if pos1 then 
				a1=pos1/2
			 	a2=pos1 mod 2
			 	indice=repere+a1-a2
				indice=repere+pos1/2-(pos1 mod 2) : chaine=chaine+chr$(13)+an(indice)
			endif
		Sheet.getCellByPosition(colonne, ligne).string=chaine
		endif
	fin1:
	  next ligne
	 next colonne
	next j


'ajustement automatique lignes et colonnes
	'feuilles planètes
	for j=num0 to num1
	Sheet = Doc.Sheets.getByName(planete(j))
		'ajustement de la largeur des colonnes
		for i=0 to 17
		Sheet.columns(i).Optimalwidth = True
		next i
		'ajustement de la hauteur des lignes
		for i=0 to 79
		Sheet.Rows(i).OptimalHeight = True
		next i
	next j
	'feuilles maisons et signes
	tr=array("maisons","signes")
	for j=0 to 1
		Sheet = Doc.Sheets.getByName(tr(j))
		'ajustement de la largeur des colonnes
		for i=0 to 7
		Sheet.columns(i).Optimalwidth = True
		next i
		'ajustement de la hauteur des lignes
		for i=0 to 12
		Sheet.Rows(i).OptimalHeight = True
		next i
	next j
	
'affichage nom et dates dans la form (si form 'visuel' active)
	if form_ok=1 then
	Sheet = Doc.Sheets.getByName(planete(num1))
		if nom_tableau="transits2" then abc="Label7"
		if nom_tableau="progressé" then abc="Label8"
		if nom_tableau="transits1" then abc="Label10"
	commande1 = feuille.getControl(abc)
		'suppression du line feed
		abc=sheet.getcellbyposition(0,0).getstring
		abc= Replace$(abc,chr$(10), " ")
	commande1.text=abc
	endif


'*******************************option de sauvegarde*************************************
	bcd=""
	choix=inputbox ("sauvegarde des feuilles au format html ? (O,N)", "confirmation","O")
	if choix<> "O" then goto fin

'copie des feuilles dans un nouveau document doc2
doc2 = StarDesktop.loadComponentFromUrl("private:factory/scalc" , "_blank",0,dimArray())
	for i = num1 to num0 step -1	
		'copie à partir de doc
		selectSheetByName(doc, planete(i))
		dispatchURL(doc,".uno:SelectAll")
		dispatchURL(doc,".uno:Copy")
		'crée une feuille si n'exite pas déjà
		If not Doc2.Sheets.hasByName(planete(i)) Then doc2.getSheets().insertNewByName(planete(i),0)
		selectSheetByName(doc2,planete(i))
			'crée la feuille suivante sinon le focus sur la feuille en cours empêche de faire le paste (copie vide)
			if i-1 < num0 then
			doc2.getSheets().insertNewByName("vide",0)
			else
			doc2.getSheets().insertNewByName(planete(i-1),0)
			endif
		'copie dans doc2
		dispatchURL(doc2,".uno:Paste")
	next i

'sauvegarde en html
	Sheet = Doc.Sheets.getByName(nom_tableau)
	'nom thème + type de transits si progressés
	abc= Sheet.getCellByPosition(7,0).getstring
		'suppression du line feed
		abc= Replace$(abc,chr$(10), "-")
	'nom du fichier
	abc= abc & "_transits par planètes "
	'filtres
	filterArgs(0).Name  = "FilterName"
  	filterArgs(0).Value = "HTML (StarCalc)" '"Text - txt - csv (StarCalc)"
	'filterArgs(1).Name  = "FilterOptions"
  	'field sep(44 - comma), txt delim (34 - dblquo), charset (0 = system, 76 - utf8), first line (1 or 2)
	'filterArgs(1).Value = "44,34,76,1"
	 bcd="file:///"  & curdir  & "/" &  abc & anneemin & "-" & anneemax & ".html"
'	ThisComponent.storeAsURL(bcd, filterArgs)
	doc2.storeAsURL(bcd, filterArgs)
	doc2.close(true)

fin:
'affichage 1ère feuille (supprimé car message d'erreur "propriété non trouvée sheets"
'Doc=thiscomponent
'	Sheet = Doc.Sheets.getByName("maisons")
'	Controller = Doc.CurrentController
'	controller.setActiveSheet(sheet)

'sactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(true)

if bcd <>"" then msgbox "terminé" & chr$(13) & bcd
end sub


'Author: Stephan Wunderlich [stephan.wunderlich@sun.com]
Sub selectSheetByName(oDoc, sheetName)
oDoc.getCurrentController.select(oDoc.getSheets().getByName(sheetName))
End Sub

Sub dispatchURL(oDoc, aURL)
Dim noProps()
Dim URL As New com.sun.star.util.URL
Dim frame
Dim transf
Dim disp
frame = oDoc.getCurrentController().getFrame()
URL.Complete = aURL
transf = createUnoService("com.sun.star.util.URLTransformer")
transf.parseStrict(URL)
disp = frame.queryDispatch(URL, "", _
com.sun.star.frame.FrameSearchFlag.SELF _
OR com.sun.star.frame.FrameSearchFlag.CHILDREN)
disp.dispatch(URL, noProps())
End Sub



sub tableau_par_annee
dim doc2 as object
Dim aBorder as New com.sun.star.table.BorderLine
Dim oBorder as New com.sun.star.table.TableBorder
Dim filterArgs(1) as new com.sun.star.beans.PropertyValue
dim anneemin, anneemax as string
dim annee%
dim finligne as string
dim jour as string
dim mois as integer
dim nom_tableau as string
dim phase_generique as string
dim symbole_transitee, symbole_aspect_transit, symbole_signe as string
dim symbole_ok as integer
dim transitante, transitee, phase_transit, aspect_transit as string
dim ind0%,ind1%
dim opt(2) as string
dim tr(2) as string

call definitions
Doc=thiscomponent

'vérifie la présence de feuilles de transits
tr=array("Uranus","Mercure","Soleil")
opt=array("- les mondiales (1) "," - les personnelles (2) "," - le thème progressé (3)")
abc="tableaux par années pour "
car=0
	for i= 0 to 2
	If Doc.Sheets.hasByName(tr(i)) Then
		abc=abc & opt(i)
		if car = 0 then car=i+1
		endif
	next i
if car=0 then msgbox "pas de feuilles planètes, exécuter : tableaux par planètes" : exit sub


'choix planètes
ref0:
car=val(inputbox (abc, "tableau annuels par mois", car))
if car=0 then exit sub
if car < 1 or car > 3 then goto ref0
	if car=2 then
	num0=2
	num1=4
	nom_tableau="transits1"
	elseif car=1 then
	num0=5
	num1=11
	nom_tableau="transits2"
	else
	num0=0
	num1=6
	nom_tableau="progressé"
	endif

'vérification présence feuilles transits et planètes	
If not Doc.Sheets.hasByName(nom_tableau) then msgbox "abandon, manque feuille " & nom_tableau : exit sub	
for i=num0 to num1
If not Doc.Sheets.hasByName(planete(i)) then msgbox "abandon, manque feuille " & planete(i) : exit sub
next i


'récupération anneemin et anneemax
	Sheet = Doc.Sheets.getByName(planete(num1))
	abc=Sheet.getCellByPosition(0,0).getstring
	if instr(1,abc,"-") then
	anneemin=mid$(abc,instr(1,abc,"-")-5,4)
	anneemax=mid$(abc,instr(1,abc,"-") +2,4)
	else
	anneemin=mid$(abc,len(abc)-3,4)
	anneemax=anneemin
	endif

'choix des années de début et fin
if anneemin <> anneemax then
	ref1:
	choix=inputbox ("entrer année de départ" , "tableau transits par années " & anneemin & " - " & anneemax, anneemin)
	if choix="" then exit sub
	if choix < anneemin or choix > anneemax then goto ref1
	anneemin = choix
	ref2:
	choix=inputbox ("entrer année de fin" , "tableau transits par années " & anneemin & " - " & anneemax, anneemax)
	if choix="" then exit sub
	if choix < anneemin or choix > anneemax then goto ref2
	anneemax = choix
endif

'choix utilisation symboles ou non
choix=inputbox("utilisation de symboles pour les aspects et planètes (O, N)  ?", "confirmation","O")
if choix="" then exit sub
	abc="aspects et planètes écrits en "
	if choix = "O" or choix="o" then
	symbole_ok =1 : abc=abc & "symboles"
	else
	symbole_ok=0 : abc=abc & "texte"
	endif

'confirmation
choix=inputbox("confirmer la création des feuilles annuelles (par mois) pour les années " & anneemin & " - " & anneemax & " (O,N) ?", abc,"O")
if choix <> "O" then exit sub

'désactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(false)

'limites basses et hautes de la barre de progression (si form 'visuel' active)
	if form_ok=1 then
	'pour écriture feuilles
	commande1 = feuille.getControl("ProgressBar1")
	commande1.setrange(num0,num1)
	commande1.setvalue(num0)
	'pour création feuilles
	commande2= feuille.getControl("ProgressBar2")
	commande2.setrange(val(anneemin),val(anneemax))
	commande2.visible=true
	feuille.getControl("Label3").text="création feuilles années"
	endif
	
'****ne pas faire de copie de feuilles à partir de la 1ère à cause des phases génériques !!!!****
'suppression si existent et recréation des feuilles annuelles
for m = val(anneemin) to val(anneemax)
	'affichage barre de progression (si form 'visuel' active)
	if form_ok=1 then commande2.setvalue(m)
	If Doc.Sheets.hasByName(m) Then Doc.Sheets.RemoveByName(m)
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
 	Doc.Sheets.insertByName(m,Sheet)

'configuration feuille	
Sheet = Doc.Sheets.getByName(m)
 	'alignement du texte : centre H et V + taille caractères
	sheet.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	sheet.VertJustify = com.sun.star.table.CellVertJustify.CENTER
 	sheet.charheight=6
 	'couleurs bleu clair du cadre
 		'ligne du haut
		sheet.getCellRangeByPosition(0,0,num1-num0+1,0).CellBackColor = RGB(153,204,255) 'bleu clair
		'1ère colonne
		sheet.getCellRangeByPosition(0,0,0,15).CellBackColor = RGB(153,204,255) 'bleu clair
		'dernière colonne
		sheet.getCellRangeByPosition(num1-num0+2,0,num1-num0+2,15).CellBackColor = RGB(153,204,255) 'bleu clair
		
 	'lecture nom feuille transits
Sheet = Doc.Sheets.getByName(nom_tableau)
 	abc= Sheet.getCellByPosition(7,0).getstring
 	'écriture "nom + année" dans 1ère cellule feuille année
Sheet = Doc.Sheets.getByName(m)
  	Cell = Sheet.getCellByPosition(0,0) : cell.string=abc & chr$(13) & str(m)
 	'couleur blanche pour 1ère cellule
 	Cell.CellBackColor = RGB(255,255,255) 'blanc
 	'en-tête ligne maison
 	Sheet.getCellByPosition(0,1).string="Maison"
 	'en-tête ligne maison
 	Sheet.getCellByPosition(0,2).string="Signe"
 	'en-tête ligne phase générique
 	Sheet.getCellByPosition(0,3).string="phase générique"
 
 
  car=1
  'planètes
  for i=num0 to num1
  
  	'phases génériques
Sheet = Doc.Sheets.getByName("psychogenèse")
	'tableau progressé présent dans la feuille thème ?
	if instr(1,Sheet.getCellByPosition(0,2).getstring,"NL") then
		phase_generique=""
		'recherche ligne avec même date = m
		for j=2 to 111 '53 to 150
			'ligne trouvée
			if instr(1,Sheet.getCellByPosition(0,j).getstring,m) then
			'recherche si phase générique présente
			car1=instr(1,Sheet.getCellByPosition(i+1,j).getstring,chr$(966))
			'lecture phase générique
			if car1 then phase_generique=mid$(Sheet.getCellByPosition(i+1,j).getstring,car1)
			exit for
			endif
		next j
	endif
			
	'en-têtes planète	
Sheet = Doc.Sheets.getByName(m)
	Cell = Sheet.getCellByPosition(car,0)
		'texte ou symboles
		if symbole_ok=0 then
		cell.string=planete(i)
		else
		cell.CharFontName="Zodiac S"
		cell.string=chr$(65+i) '65=A, de A à J Soleil à Pluton
		if i =10 or i=11 then cell.string=chr$(65+i+2) 'NN et Lilith, M et N
		endif
	'en-tête phase générique	
	Sheet.getCellByPosition(car,3).string=phase_generique
		
  car=car+1
  next i
  
  	'en-têtes lignes (mois)
	for i = 4 to 15
	Cell = Sheet.getCellByPosition(0,i) : cell.string =format("1/"&i-3,"mmmm")' mois en lettres
	Cell = Sheet.getCellByPosition(num1-num0+2,i) : cell.string =format("1/"&i-3,"mmmm")
	next i
	'bordure cellules
	call cadre_cellules(m,8,15) 'feuille, dernière colonne, dernière ligne

next m



'************début*********
if form_ok=1 then feuille.getControl("Label3").text="écriture feuilles années"
'lecture des données sur feuilles des planètes- écriture des données sur feuille année
for i=num0 to num1
	'affichage barre de progression (si form 'visuel' active)
	if form_ok=1 then commande1.setvalue(i)
Sheet = Doc.Sheets.getByName(planete(i))
	'colonnes planètes
	for j=1 to 16
	car=1
		'lignes aspects de 0 à 79
		do until car > 77
			transitante=planete(i)
			aspect_transit=Sheet.getCellByPosition(0,car).getstring
			transitee=Sheet.getCellByPosition(j,0).getstring
			'élimination de certains transits
			select case transitante
				case "Jupiter"
				if instr(1,aspect_transit,"conjonction")=0  and instr(1,aspect_transit,"opposition")=0  and instr(1,aspect_transit,"carré")=0  and instr(1,aspect_transit,"trigone")=0 then goto finboucle
				if instr(1,aspect_transit,"semi") or instr(1,aspect_transit,"sesqui") then goto finboucle
				case "NN"
				if instr(1,aspect_transit,"conjonction")=0  and instr(1,aspect_transit,"opposition")=0  and instr(1,aspect_transit,"carré")=0  then goto finboucle
				if instr(1,aspect_transit,"semi") or instr(1,aspect_transit,"sesqui") then goto finboucle
				case "Lilith"
				if instr(1,aspect_transit,"conjonction")=0 then goto finboucle
			end select
			select case transitee
				case "FC","DS"
				goto finboucle
					if nom_tableau="progressé" then goto suite
				case "AS","MC","NN"
				if instr(1,aspect_transit,"conjonction")=0  and instr(1,aspect_transit,"opposition")=0  and instr(1,aspect_transit,"carré")=0 then goto finboucle
				if instr(1,aspect_transit,"semi") or instr(1,aspect_transit,"sesqui") then goto finboucle
			end select
suite:				
			'lecture des 4 lignes de "phases" d'un même aspect
			for k=car to car+3
			abc= Sheet.getCellByPosition(j,k).getstring
				'données trouvées (une ou plusieurs dates)
				if abc <> "" then
				ind0=1
					'phase ?
					phase_transit=""
					if k <> car then phase_transit=" (" & Sheet.getCellByPosition(0,k).getstring & ")"
					'extraction année, mois, jour (plusieurs fois éventuellement)
					do
						ind1=instr(ind0,abc,chr$(10))
						'date (si ind1, fait partie d'une série)
						if ind1 then bcd=mid$(abc,ind0,ind1-ind0) else bcd=right(abc,10)
						annee=year(bcd)
						
						'*******année trouvée dans la plage cherchée*********
						if annee >= val(anneemin) and annee <= val(anneemax) then 
							mois=month(bcd)
							if mois < 1 or mois > 12 then msgbox ("erreur mois " & str(mois) & transitante & aspect_transit & transitee & abc)
							jour= "  " & day(bcd)
																	
							'écriture dans la feuille correspondant à l'année
							Sheet = Doc.Sheets.getByName(annee)
							Cell = Sheet.getCellByPosition(i-num0+1,mois+3)
							finligne=""
							if cell.string <> "" then finligne=chr$(13)
							
						'écriture données avec symboles ou texte
							'aspects et planètes écrits en symboles	avec la police "Zodiac S" 
												 '9 aspects (conjonction à opposition) = m à u
												 '12 signes (bélier à poissons) = a à l (chr$(97 à 108)
												 '10 planètes (soleil à pluton) = A à J
												 '4 compléments AS=K, MC=L, NN=M, Lilith=N
												 'flèche aspect croissant = +  décroissant = ,
							if symbole_ok=1 then	 
								'planètes
								symbole_transitee=""
								for m=0 to 9
									if transitee=planete(m) then
									symbole_transitee=chr$(65+m) '65=A, de A à J Soleil à Pluton
									goto suite1
									endif
								next m
									if transitee = "NN" then symbole_transitee="M"
									if transitee = "Lilith" then symbole_transitee="N"
									if transitee = "AS" then symbole_transitee="K"
									if transitee = "FC" then symbole_transitee="L,"
									if transitee = "DS" then symbole_transitee="K,"
									if transitee = "MC" then symbole_transitee="L"
									
								'aspects		
							suite1:
								'aspects croissants conjonction à opposition
								for m=0 to 8
									if aspect_transit=aspect(m) then
									symbole_aspect_transit=chr$(109+m) '109=m, de m à u conjonction à opposition
									if m > 0 and m < 8 then symbole_aspect_transit=symbole_aspect_transit & "+" 'aspects croissants sauf conjonction et opposition
									goto suite2
									endif
								next m
								'aspects décroissants
								for m=9 to 15
									if aspect_transit=aspect(m) then
									symbole_aspect_transit=chr$(109+m-2*(m-8)) '109=m, de m à u conjonction à opposition
									symbole_aspect_transit=symbole_aspect_transit & "," 'aspects décroissants
									goto suite2
									endif
								next m
								msgbox "pas de symbole trouvé" & transitante & aspect_transit & transitee
								
							suite2:	
								cell.CharFontName="Zodiac S"
								'écriture aspects et planètes en symboles
								cell.string= cell.string  & finligne & symbole_aspect_transit & symbole_transitee & jour						
							else
								'écriture aspects et planètes en clair	
								cell.string= cell.string  & finligne & aspect_transit & " " & transitee & phase_transit & jour
							endif
							
							'couleur de fond éventuellement
							if instr(1,phase_transit,"1") then Cell.CellBackColor = RGB(167, 236, 106) 'vert
							if instr(1,phase_transit,"2") then Cell.CellBackColor = RGB(255, 213, 6) 'orange
							if instr(1,phase_transit,"3") then Cell.CellBackColor = RGB(255, 153, 153) 'rouge
							'if transitee ="AS" or transitee = "FC" or transitee = "DS" or transitee ="MC" then Cell.CellBackColor = RGB(255,255,204) 'jaune	clair
						endif				
					'si ind1, recherche des autres dates dans la même cellule
					if ind1=0 then exit do
					ind0=ind1+1
					loop
				Sheet = Doc.Sheets.getByName(planete(i))
				endif
			next k
finboucle:
		'aspect suivant
		car=car+5
		loop
	next j
next i	

'lecture des données  feuilles maisons et signes - écriture dans feuille année (données uniquement en progressé)
	tr=array("maisons","signes")
for k=0 to 1
	'colonnes planètes transitantes
	for i=1 to 7
		'lignes maisons (k=0) ou signes (k=1)
		for j= 1 to 12
		Sheet = Doc.Sheets.getByName(tr(k))
		abc=Sheet.getCellByPosition(i,j).getstring
			if abc<>"" then
			ind0=1
				do
				ind1=instr(ind0,abc,chr$(10))
					'date (si ind1, plusieurs dates dans la cellule)
					if ind1 then bcd=mid$(abc,ind0,ind1-ind0) else bcd=right(abc,10)
				annee=year(bcd)	
					if annee >= val(anneemin) and annee <= val(anneemax) then 
						mois=month(bcd)
						jour= "  " & day(bcd)
							'écriture dans la feuille correspondant à l'année, ligne 1 pour maisons, ligne 2 pour signes
							Sheet = Doc.Sheets.getByName(annee)
							Cell = Sheet.getCellByPosition(i,k+1)
								finligne=""
								if cell.string <> "" then finligne=chr$(13)
								'symboles ou en clair ?
								if symbole_ok=1 then
									cell.CharFontName="Zodiac S"
									symbole_aspect_transit=chr$(109)
									symbole_signe=chr$(96+j)
								else
									symbole_aspect_transit="conjonction"
									symbole_signe=signe(j-1)
								endif
							 'écriture
							'cuspide
							if k=0 then cell.string= cell.string  & finligne & symbole_aspect_transit & j & jour  &"/" & mois	
							'signe
							if k=1 then	cell.string= cell.string & finligne & symbole_signe & jour &"/" & mois
					endif
				'si ind1, recherche des autres dates dans la même cellule
				if ind1=0 then exit do
				ind0=ind1+1
				loop
			endif
		next j
	next i
next k
			
'ajustement automatique lignes et colonnes
for m= val(anneemin) to val(anneemax)
	Sheet = Doc.Sheets.getByName(m)
	'ajustement de la largeur des colonnes
	for i=0 to num1 - num0 +2
	Sheet.columns(i).Optimalwidth = True
	next i
	'ajustement de la hauteur des lignes
	for i=0 to 13
	Sheet.Rows(i).OptimalHeight = True
	next i
next m

'option de sauvegarde
	bcd=""
	choix=inputbox ("sauvegarde des feuilles au format html ? (O,N)", "confirmation","O")
	if choix<> "O" then goto fin

'copie des feuilles dans un nouveau document doc2
	doc2 = StarDesktop.loadComponentFromUrl("private:factory/scalc" , "_blank",0,dimArray())
	'copie feuilles années
	for i=val(anneemax) to val(anneemin) step-1
		selectSheetByName(doc, i)
		dispatchURL(doc,".uno:SelectAll")
		dispatchURL(doc,".uno:Copy")
		'crée une feuille si n'existe pas déjà
		If not Doc2.Sheets.hasByName(i) Then doc2.getSheets().insertNewByName(i,0)
		selectSheetByName(doc2, i)
		'crée une feuille vide sinon le focus sur la feuille en cours empêche de faire le paste (copie vide)
				if i-1 < val(anneemin) then
				doc2.getSheets().insertNewByName("vide",0)
				else
				doc2.getSheets().insertNewByName(i-1,0)
				endif
			'copie dans doc2
		dispatchURL(doc2,".uno:Paste")
	next i
	'copie feuilles maisons et signes si progressé
	tr=array("maisons","signes")
	for i=0 to 1
		Sheet = Doc.Sheets.getByName(tr(i))
		if instr(1,Sheet.getCellByPosition(0,0).getstring,"progressé") then
			selectSheetByName(doc, tr(i))
			dispatchURL(doc,".uno:SelectAll")
			dispatchURL(doc,".uno:Copy")
			If not Doc2.Sheets.hasByName(tr(i)) Then doc2.getSheets().insertNewByName(tr(i),0)
			selectSheetByName(doc2,tr(i))
			'crée une feuille vide sinon le focus sur la feuille en cours empêche de faire le paste (copie vide)
			doc2.getSheets().insertNewByName("vide" & i,0)
			dispatchURL(doc2,".uno:Paste")
		endif
	next i
	
'sauvegarde en html
	Sheet = Doc.Sheets.getByName(nom_tableau)
	'nom thème + type de transits si progressés
	abc= Sheet.getCellByPosition(7,0).getstring
		'suppression du line feed
		abc= Replace$(abc,chr$(10), "-")
	'nom du fichier
	abc= abc & "_transits par années "
	'filtres
	filterArgs(0).Name  = "FilterName"
  	filterArgs(0).Value = "HTML (StarCalc)" '"Text - txt - csv (StarCalc)"
	'filterArgs(1).Name  = "FilterOptions"
  	'field sep(44 - comma), txt delim (34 - dblquo), charset (0 = system, 76 - utf8), first line (1 or 2)
	'filterArgs(1).Value = "44,34,76,1"
	bcd="file:///"  & curdir  & "/" &  abc & anneemin & "-" & anneemax & ".html"
	'ThisComponent.storeAsURL(bcd, filterArgs)
	doc2.storeAsURL(bcd, filterArgs)
	doc2.close(true)
	
'option de suppression des feuilles
	choix=inputbox ("suppression des feuilles années dans le tableur ? (O,N)", "confirmation suppression feuilles années","O")
	if choix<> "O" then goto fin
	for i=val(anneemin) to val(anneemax)
	If Doc.Sheets.hasByName(i) Then Doc.Sheets.RemoveByName(i)
	next i
	goto fin2
fin:		

'mise au 1er plan de la feuille
Doc=thiscomponent
	Sheet = Doc.Sheets.getByName(anneemin)
	Controller = Doc.CurrentController
	controller.setActiveSheet(sheet)

'sactive boutons tableaux, thème, aspects, etc.
call actions_boutons2(true)
	
fin2:	
if bcd <> "" then msgbox "terminé" & chr$(13) & bcd
end sub



sub zodiaque (optional choix0,optional choix as integer) 'choice0 pour éliminer "fausse valeur" due à l'appel à partir de la form visuel

Dim cercle As Object
Dim Page As Object
dim ligne(50) as object 'nombre max d'aspects
dim carte(55) as object
Dim Point As New com.sun.star.awt.Point
Dim Size As New com.sun.star.awt.Size
Dim Gradient As New com.sun.star.awt.Gradient 
Dim x2 as New com.sun.star.util.Date
Dim y2 As New com.sun.star.util.Time
'Dim x2 As New com.sun.star.awt.XDateField

dim aa
dim ag_deb%, ag_fin%
dim an1 as long
dim annee%

dim annee_naissance
dim anneemin as long, anneemax as long
dim anniversaire as string
dim astre
dim bb
dim curseur1, curseur1ref, curseur2, curseur2ref, curseur3, curseur3ref, curseur4, curseur4ref
dim date_anniv
dim date_calcul
dim date_janv as long, datex as long, date_jour as long
dim date_string as string
dim ecart_heure as double
dim ecart_rang
dim gap
dim h1,h2,d2,l1,l2
Dim Height As Long
dim heure as double
dim heure_naissance as double, heure_naissance_ref as double 'pas long sinon =0 !
dim heure_string as string
dim heure_naissance_string as string
dim hh%, hh_naissance%
dim inc
dim index1
dim jour%
dim longitude0, longitude1
dim mondialok%
dim maisonsok%
dim mm%, mm_naissance%
dim mois%
dim nom,num
dim offset
dim option_indice%, option2_indice%, option3_indice%
dim orbe_as_fc 'array pour orbes AS à MC
dim choix%
dim choixref
dim rang as long, rang_mondial as long
dim sinus
dim su_deb%, su_fin%
dim theme as string
dim tol1 as double, tol2 as double
dim tol1min as double, tol2min as double
dim val1
Dim Width As Long
Dim x As Date
dim y 'as time


dim coul(0 to 12) as long
dim ecart(15)
dim etiquette(1 to 5) as string
dim etiquette2(1 to 3) as string
dim etiquette3(1 to 2) as string
dim longitudes(15) as double
dim maison(1 to 12) as double
dim natal(27) as double
dim options(1 to 5) as string
dim options2(1 to 3) as string
dim options3(1 to 2) as string
dim tolerance(31)


		
oFA = createUnoService( "com.sun.star.sheet.FunctionAccess" )
call definitions
message=""

'aspects et planètes écrits en symboles	avec la police "Zodiac S" 
 '9 aspects (conjonction à opposition) = m à u
'12 signes (bélier à poissons) = a à l
 '10 planètes (soleil à pluton) = A à J
'4 compléments AS=K, MC=L, NN=M, Lilith=N
 'flèche aspect croissant = +  décroissant = ,
 Doc = thisComponent
if not Doc.Sheets.hasByName("éphémérides") then msgbox "manque feuille éphémérides, lancer 'calcul thème'" : exit sub

'véification nom du thème
	Sheet = Doc.Sheets.getByName("éphémérides")
	theme=sheet.getcellbyposition(19,0).getstring
	if theme="" then msgbox "pas de nom de thème page éphémérides, lancer 'calcul thème'" : exit sub 

'récupère le nombre de lignes utilisées dans la feuille éphémérides
	Curs = Sheet.createCursor
	Curs.gotoEndOfUsedArea(True)
	lastrow = Curs.Rows.Count -1
	anneemin=val(right(sheet.getcellbyposition(0,17).getstring,4))
	anneemax=val(right(sheet.getcellbyposition(0,lastrow).getstring,4))
	
'trop compliqué à gérer à cause du natal (dates non initialisées correctement) ! si feuille visuel on passe la question et on affiche le natal
	if form_ok=1 then choix=1 : goto ref1
		
'choix zodiaque natal ou  progressé, si non automatique (si lancé depuis 'calcul du thème' avec choix=1)
'If choix0 <>1 then 
if choix <1 or choix >2 then
ref0:
	choix=val(inputbox("création d'un zodiaque natal (1), progressé/progressé (2), progressé/natal (3), mondial (4) ou mondial/natal (5) ?", "choix zodiaque natal, progressé ou mondial", 1))
	if choix=0 then exit sub
	if choix <1 or choix > 5 then goto ref0 
endif
'endif

ref1:
		'1ère année des éphémérides
			abc=Sheet.getCellByPosition(0, 17).getstring
			if abc="" then message="manque éphémérides, options désactivées"
			if abc="" and choix >1 then msgbox message : exit sub
			an1=datevalue(Sheet.getCellByPosition(0, 17).getstring)
		'date de naissance
			date_anniv=Sheet.getCellByPosition(22, 1).getstring
			if date_anniv=""  then msgbox "impossible, pas de date de naissance définie" : exit sub
		'valeur numérique de la date de naissance
			date_naissance=datevalue(date_anniv)
			if date_naissance<0 then message="date de naissance < 1900, options désactivées"
			if date_naissance<0 and choix >1 then msgbox message : exit sub
		'année de naissance
			annee_naissance=year(date_anniv)
		'date anniversaire jj/mm (ex "30/10")
			anniversaire=mid$(date_anniv,1,6)
		'date du jour
			'format numérique
			date_jour=datevalue(now)
			'format string
			date_string=date(now)
		'jour
			jour=day(now)
		'mois en cours
			mois=month(now)
		'année en cours
			annee=year(now)
		'valeur numérique de la date  anniversaire de l'année en cours
			datex=datevalue(Replace$(anniversaire  & str(annee)," ", ""))
		'heure de naissance
			abc=Sheet.getCellByPosition(22, 2).getstring
			'format string
			heure_naissance_string=abc
			'format numérique
			heure_naissance=timevalue(abc)
			heure_naissance_ref=heure_naissance
		'pour le progressé (recalcule jours et annee_progresse)
		call calc_progresse(datex,date_jour,annee,anniversaire)
			'tolérances par défaut : orbe année
			tol1=1 : tol2=1
		'pour le mondial
			'ligne éphémérides du jour
			rang_mondial=date_jour-an1+17
			'heure actuelle
				abc=time(now)
				'coefficient de 0 à 1 (midi=0,5, minuit=1)
				heure=timevalue(abc) 
				'format string	
				heure_string=mid$(abc,1,5)

		choixref=choix
		'affichage maisons et aujourd'hui actifs, anniversaire inactif
		maisonsok=1
		'thème mondial non actif par défaut
		mondialok=0
		'tolérances mini progressé et mondial/natal
		tol1min=1/360 : tol2min=1/3
		clic_sortie=0
'endif

	
'supprime et recrée feuille "zodiaque"
	if Doc.Sheets.hasByName("zodiaque") Then Doc.Sheets.RemoveByName("zodiaque")
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
	Doc.Sheets.insertByName("zodiaque", Sheet)
	
'récupère index feuille zodiaque
	For i = 0 to Doc.Sheets.Count-1
		if Doc.Sheets(i).Name="zodiaque" then index1=i
	next i
	
'déclaration lignes	
	for i= 0 to 50
 		ligne(i) = Doc.createInstance("com.sun.star.drawing.LineShape")
	next i
	
'déclaration cercles
	for i=0 to 55
		carte(i)=Doc.createInstance("com.sun.star.drawing.EllipseShape")
	next i
	'couleur de fond jaune
	Sheet = Doc.Sheets.getByName("zodiaque")
	Sheet.getCellrangeByPosition(0,0,5,21).cellBackColor = RGB(255, 255, 201)
	
'définit la page pour les graphiques	
Page = Doc.drawPages(index1)


'label + noms boutons radios (= type de zodiaque)
	aa=array("natal","progressé/progressé","progressé/natal","carte du ciel","transits/natal")
	bb=array("option1","option2","option3","option4","option5")
	for i= 1 to 5
		etiquette(i)=aa(i-1)
		options(i)=bb(i-1)
	next i
	
'label + noms boutons radios (= aujourd'hui, anniversaire, variable)
	aa=array("maintenant","anniversaire","variable")
	bb=array("Bouton1","Bouton2","Bouton3")
	for i= 1 to 3
		etiquette2(i)=aa(i-1)
		options2(i)=bb(i-1)
	next i

'label + noms boutons radios tolerance
	aa=array("tol. max","tol.min")
	bb=array("tol1","tol2")
	for i= 1 to 2
		etiquette3(i)=aa(i-1)
		options3(i)=bb(i-1)
	next i
		
'activation checkbox, boutons radio, date, heure (si form 'visuel' active)
	if form_ok=1 then 
	
		'active bouton sortie
			feuille.getControl("sortie").enable=true
		'désactive boutons tableaux, thème, aspects, etc.
			call actions_boutons2(false)
				
		'cache date,heure,heure de naissance,checkbox mondial, 3 boutons "maintenant, anniversaire et variable" et 2 boutons tolérance
			call actions_boutons(false,false,false,false,false,false)	
											
		'affiche et active checkbox maisons	
			feuille.getControl("CheckBox1").visible=true
			feuille.getControl("CheckBox1").state=1
		
		'désactive checkbox mondial
			feuille.getControl("CheckBox2").state=0		
		
		'boutons radio des options natal,progressé, etc.
		for i=1 to 5
			feuille.getControl(options(i)).label=etiquette(i)
			if choix=i then feuille.getControl(options(i)).state=1 : option_indice=i
			feuille.getControl(options(i)).visible=true
			'boutons cachés si natal ou date de naissance < 1900 sinon erreurs avec numéros de lignes éphémérides
			if date_naissance<0 and i >1 then feuille.getControl(options(i)).visible=false
		next i
		
		'boutons radio des options aujourd'hui, anniversaire, variable
		for i= 1 to 3
			feuille.getControl(options2(i)).label=etiquette2(i)
			feuille.getControl(options2(i)).state=0
		next i
			'active le bouton 'aujourd'hui'
			feuille.getControl(options2(1)).state=1 : option2_indice=1 '(1=aujourd'hui, 2=anniversaire)
				
		'boutons radio tolérance
		for i=1 to 2
			feuille.getControl(options3(i)).state=0 
			feuille.getControl(options3(i)).label=etiquette3(i)
		next i
			'active bouton max
			feuille.getControl(options3(1)).state=1 : option3_indice=1 '(1=max, 2=min)
				
		'date : commande1
		commande1=feuille.getControl("DateField1")
			'affichage année sur 4 chiffres
			commande1.model.dateformat=7
		x = Date()
		x2.Year = annee 'Year(x)
		x2.Month = mois 'Month(x)
		x2.Day =jour ' Day(x)		
		commande1.Date = x2
		'valeurs min
		x2.Year = anneemin
		x2.Month = 1
		x2.Day = 1
		commande1.min=x2
		'valeurs max
		x2.Year = anneemax
		x2.Month = 12
		x2.Day = 31
		commande1.max=x2
		
		'heure:commande2
		commande2=feuille.getControl("TimeField1")
		y = Time()
		y2.Hours = Hour(y)
		y2.Minutes = Minute(y)
		commande2.Time = y2
		hh=Hour(y)
		mm=Minute(y)
		
		'heure de naissance : commande3
		commande3=feuille.getControl("TimeField2")
		y = heure_naissance_string
		y2.Hours = Hour(y)
		y2.Minutes = Minute(y)
		commande3.Time=y2
		hh_naissance=Hour(y)
		mm_naissance=Minute(y)
		if message="" then
			feuille.getControl("Label3").text="heure de naissance"
		else
			feuille.getControl("Label3").text=message
		endif
				
	endif
	
Sheet = Doc.Sheets.getByName("éphémérides")

'récupération des orbes page éphémérides pour AS à MC
	cell=Sheet.getCellRangeByPosition(17,1,17,16)
	orbe_as_fc=cell.getdata

'couleurs pour symboles planètes et signes
	'couleur du Bélier (mauve)
	coul(0)= RGB(255, 100, 255)
	for i=1 to 12
	coul(i)= coul(i-1)+20000
	next i
	'cas particulier où la longitude=360 exactement
	coul(12)=coul(0)


				
'***************************début de la partie inteactive pour incrémentation des années***********************************
debut:	

Sheet = Doc.Sheets.getByName("éphémérides")

'planètes natales
	'Soleil à MC
	for i=0 to 15
		longitudes(i)=Sheet.getCellByPosition(i+1, 1).getvalue
		if choix >1 then
		'en double pour thème interne + positions peuvent varier avec l'heure de naissance
		rang=date_naissance-an1+17
			ecart(i)=Sheet.getCellByPosition(i+1,rang+1).getvalue-Sheet.getCellByPosition(i+1,rang).getvalue
				if ecart(i) < -300 then ecart(i)=ecart(i)+360
				if ecart(i) > 300 then ecart(i)=ecart(i)-360
		natal(i)=longitudes(i)+heure_naissance*(ecart(i))
		endif
	next i


'planètes progressées et mondiales
 	'thème progressé
	if choix =2 or choix=3 then
		'ligne analysée
		rang=date_naissance -an1 + 17 + annee_progresse -val(annee_naissance)
		'Soleil à Lilith
		for i=0 to 11
			'écart sur 1 an
			ecart(i)=Sheet.getCellByPosition(i+1,rang+1).getvalue-Sheet.getCellByPosition(i+1,rang).getvalue
				if ecart(i) < -300 then ecart(i)=ecart(i)+360
				if ecart(i) > 300 then ecart(i)=ecart(i)-360
			'offset pour heure de naissance (important sinon décalage possible de plusieurs mois si naissance après-midi par exemple)
			ecart_heure=heure_naissance*(ecart(i))
			'tolérance 1 an= + ou - 1/2 an
			tolerance(i)=abs(ecart(i)/2)
			'position à la date située à date précédente d'anniversaire + jours + 1/2 journée (=à midi)
			longitudes(i)=Sheet.getCellByPosition(i+1,rang).getvalue+ecart_heure+jours*ecart(i)/365+ecart(i)/730
				if longitudes(i) >= 360 then longitudes(i)=longitudes(i)-360
				if longitudes(i) < 0 then longitudes(i)=longitudes(i)+360
		next i
		
	'mondial
	elseif choix=4 or choix=5 then
		'ligne analysée
		rang=rang_mondial
			'limite basse éphémérides
			if rang < 17 then rang=17
			'limite haute éphémérides : reste à l'avant-dernière ligne sinon ecart() est faux (0-x)
			if rang >= lastrow then rang=lastrow-1
		'Soleil à Lilith
		for i=0 to 11
			'écart sur 24h 
			ecart(i)=Sheet.getCellByPosition(i+1,rang+1).getvalue-Sheet.getCellByPosition(i+1,rang).getvalue
				if ecart(i) < -300 then ecart(i)=ecart(i)+360
				if ecart(i) > 300 then ecart(i)=ecart(i)-360
			'offset pour heure
			ecart_heure=heure*(ecart(i))
			'tolérance 6 jours = + ou - 3 jours
			tolerance(i)=3*abs(ecart(i))
			'position à l'heure actuelle
			longitudes(i)=Sheet.getCellByPosition(i+1,rang).getvalue + ecart_heure
				if longitudes(i) >= 360 then longitudes(i)=longitudes(i)-360
				if longitudes(i) < 0 then longitudes(i)=longitudes(i)+360
		next i
	endif
	
'Maisons natales
	for i=1 to 12
		maison(i)=Sheet.getCellByPosition(21,i).getvalue
		'sauf en natal, les maisons peuvent varier avec l'heure de naissance
		if choix >1 then if maisonsok=1 then
			maison(i)=maison(i)+(heure_naissance-heure_naissance_ref)*360
				if maison(i) >= 360 then maison(i)=maison(i)-360
				if maison(i) < 0 then maison(i)=maison(i)+360
		endif
	next i
		
'Maisons mondiales calculées à partir des maisons natales (très approximatif !)
	if choix=4 and maisonsok=1 then
		for i=1 to 12
			'coefficient en plus de l'augmentation de 360deg./24h, d'augmentation journalière des maisons
			aa=array(0.72,0.79,0.89,0.95,0.91,0.81,0.72,0.79,0.89,0.95,0.91,0.81)
			'offset années + jours (pris 1 coefficient unique sinon les maisons ne sont plus en ordre croissant!)
			car1= ((date_jour-date_naissance)*0.7003) mod 360
			'offset heure (360 deg. par jour)
			car2= (heure-heure_naissance)*360
			'ajout des 2 offsets
			bb=array(car1,car2)
				for j=0 to 1
				maison(i)=maison(i) + bb(j)
					if maison(i) >= 360 then maison(i)=maison(i)-360
					if maison(i) < 0 then maison(i)=maison(i)+360
				next j
		next i
	endif
	
'Maisons progressées calculées à partir des maisons natales
	'progressé/progressé
	if choix=2 and maisonsok=1 then 
		'coefficient en plus de l'augmentation de 360deg./24h, d'augmentation journalière des maisons = annuelle pour le progressé (approximation !)
		aa=array(0.72,0.79,0.89,0.95,0.91,0.81,0.72,0.79,0.89,0.95,0.91,0.81)
		for i=1 to 12
			'ajoute offset pour années/année de naissance + jours de l'année en cours
			maison(i)=maison(i) + aa(i-1)*(annee_progresse -val(annee_naissance)) + jours*aa(i-1)/365 
				if maison(i) > 360 then maison(i)=maison(i)-360
				if maison(i) < 0 then maison(i)=maison(i)+360
		next i
	endif
	
'ajoute les axes pour progressé et mondial = maisons équivalentes (AS=maison1, FC=maison4,DS=maison7, MC=maison10)
if choix=2 or choix=3 or choix=4 then longitudes(12)=maison(1) : longitudes(13)=maison(4) : longitudes(14)=maison(7) : longitudes(15)=maison(10)
if choix=3 or choix=5 then natal(12)=maison(1) : natal(13)=maison(4) : natal(14)=maison(7) : natal(15)=maison(10)
	
if maison(1)=0 then message="  (pas de maisons définies feuille éphémérides)"


'écriture des longitudes feuille zodiaque
	Sheet = Doc.Sheets.getByName("zodiaque")		
	sheet.charheight=6
	'en-tête colonne G
	Sheet.getCellByPosition(6,0).string="externe"
	'en-tête colonne H
	if choix =3 or choix=5 then Sheet.getCellByPosition(7,0).string="interne"
	if choix <>3 and choix <> 5 then Sheet.getCellByPosition(7,0).string="longitudes"
	'en-tête colonne I
	Sheet.getCellByPosition(8,0).string="Maisons"
	'planètes
	for i=0 to 15
		'1ère colonne F : nom des planètes
		Sheet.getCellByPosition(5,i+1).string=planete(i)
		'2ème colonne G
			' thème externe : positions en degrés minutes + signe
			coeff1=int(longitudes(i)/30)*30 : coeff2=int(longitudes(i)) mod 30
			abc=str(coeff2) & chr$(176) & str(int(60*(longitudes(i)-coeff1-coeff2))) & "' " & signe(coeff1/30) 'ex  12° 23' Lion
			Sheet.getCellByPosition(6,i+1).string=abc
		'3ème colonne H
			'longitudes détaillées du thème externe
			if choix <> 3 and choix <> 5 then Sheet.getCellByPosition(7,i+1).string=longitudes(i)
			'ou thème interne (natal) : positions en degrés minutes + signe
			if choix=3 or choix=5 then 
				coeff1=int(natal(i)/30)*30 : coeff2=int(natal(i)) mod 30
				abc=str(coeff2) & chr$(176) & str(int(60*(natal(i)-coeff1-coeff2))) & "' " & signe(coeff1/30) 'ex  12° 23' Lion
				Sheet.getCellByPosition(7,i+1).string=abc
			endif
	next i
	'maisons
	for i=1 to 12
		coeff1=int(maison(i)/30)*30 : coeff2=int(maison(i)) mod 30
		abc=str(coeff2) & chr$(176) & str(int(60*(maison(i)-coeff1-coeff2))) & "' " & signe(coeff1/30) 'ex  12° 23' Lion
		'case décochée, pas de maisons
		if maisonsok=0  then abc=""
		'écriture
		Sheet.getCellByPosition(8,i).string=abc
		Sheet.getCellByPosition(9,i).value=i
	next i
	
	'largeur colonnes F-I
	for i=5 to 9
	Sheet.columns(i).Optimalwidth = True
	next i


Sheet = Doc.Sheets.getByName("éphémérides")	

'offset >0 ou <0 entre ASC et 180 deg.pour ajuster planètes, signes et maisons
	offset=0
	if maison(1)  then if maisonsok=1 then offset=180-maison(1)


'écriture nom du thème
	abc=""
	'type de maisons si actives	
	if maisonsok=1 then
		aa=array(" -  Maisons progressées", " -  Maisons natales"," Maisons approximatives !", " -  Maisons natales - ")
		for i=2 to 5
			if choix=i then abc=aa(i-2)
		next i
	endif
	'écriture
	select case choix
	case 1
		nom=theme & " - natal - " & Sheet.getCellByPosition(22, 0).getstring & message
	case 2
		nom=theme & "  - progressé/progressé - le " & date_string & abc & " (" & heure_naissance_string & ")" & message
	case 3
		nom=theme & " - progressé/natal - le " & date_string & abc & " (" & heure_naissance_string & ")" & message
	case 4
		if mondialok=0 then
			bcd="carte du ciel - le "
		else
			bcd="thème mondial - le "
		endif
		nom= bcd & date_string & " à " & heure_string 
		nom=nom & abc & message
	case 5
		nom= "carte du ciel - le " & date_string & " à " & heure_string 
		nom=nom & " - transits sur thème natal : " & abc & theme & " (" & heure_naissance_string & ")" & message
	end select

	'écriture nom et ajustement hauteur 1ère ligne	
	Sheet = Doc.Sheets.getByName("zodiaque")
	cell=sheet.getcellbyposition(0,0)
	cell.string=nom	
	Sheet.Rows(0).OptimalHeight = True

	
'**********************************début tracés*********************************************
	
'effacement des cercles internes (remove ne marche pas)
	for i= 16 to 31
	carte(i).visible=false
	next i
		
'cercles pour symboles planètes (arcs)
	'planètes (0-15 : thème externe, 16-31 thème interne)
	for i=0 to 31
		'pas de thème interne sauf progressé/natal et mondial/natal
			if i >15 and choix <> 3 and choix <> 5 then exit for
			
		'position planète
			'thème externe : planètes natales, progressées ou mondiales
			if i <=15 then longitude=longitudes(i)+offset
			'thème interne : planètes natales
			if i >15 then longitude=natal(i-16)+offset
			if longitude <0 then longitude=longitude+360
			if longitude >360 then longitude=longitude-360
			
		'coordonnées x,y
			'thème externe planètes: cercle éventuellement agrandi pour éviter la superposition des planètes
			if i < 12 then
				car=0
				'sépare les symboles si superposition
				for j=i-1 to 0 step -1
					if longitudes(j)+offset > longitude-4 and longitudes(j)+offset < longitude+4 then car=car+1
				next j
				Point.x = 2800-(car*200)
				Point.y = 1800-(car*200) 
				Size.Width = 6400+(car*400)
				Size.Height = 6400+(car*400)
			'thème externe : axes
			elseif i >11 and i <=15 then
				Point.x = 1600
				Point.y = 600
				Size.Width = 8800
				Size.Height = 8800
			'thème interne progressé/natal
			elseif i >15 and choix=3 then
				car=0
				if i=16 then goto affi1
				'sépare les symboles si superposition
				for j=i-1 to 16 step -1
					if natal(j-16)+offset > longitude-4 and natal(j-16)+offset < longitude+4 then car=car+1
				next j
		affi1:
				Point.x = 4000+(car*200)
				Point.y = 3000+(car*200)
				Size.Width = 4000-(car*400)
				Size.Height = 4000-(car*400)
			'thème interne mondial/natal
			elseif i >15 and choix=5 then
				car=0
				if i=16 then goto affi2
				'sépare les symboles si superposition
				for j=i-1 to 16 step -1
					if natal(j-16)+offset > longitude-4 and natal(j-16)+offset < longitude+4 then car=car+1
				next j
		affi2:
				Point.x = 3500+(car*200) 
				Point.y = 2500+(car*200)
				Size.Width = 5000-(car*400)
				Size.Height = 5000-(car*400)
			endif
			
		'tracé
			carte(i).Position = point
			carte(i).Size = size
			carte(i).CircleKind =  com.sun.star.drawing.CircleKind.ARC
		'début et fin de l'arc : taille 2000 en natal ou ecart annuel en progressé = traînée de la planète à l'écran (traînée Lune seule est visible en progressé)
		'traînée Lune (Lune au centre) = 1 an (sauf au 'jour le jour' cad tol =1/180 et en natal ou mondial)
		if i=1 and tol1=1 and (choix=2 or choix=3) then
			carte(i).circlestartangle=longitude*100-(100*180*ecart(i)/365)
			carte(i).circleendangle=longitude*100+(100*180*ecart(i)/365)
			'carte(i).circlestartangle=longitude*100-(100*jours*ecart(i)/365)
			'carte(i).circleendangle=longitude*100+(100*(365-jours)*ecart(i)/365)
		'pas de traînée
		else
			carte(i).circlestartangle=longitude*100 
			carte(i).circleendangle=longitude*100 
		endif
	
		carte(i).visible=true
		'tracé de l'arc
		Page.add(carte(i))
	
			'codes ASCII des symboles
			select case i
				case 0 to 9 'thème externe
				abc=chr$(65+i) 'A à J
				case 16 to 25 'thème interne
				abc=chr$(49+i) 'A à J
				case 10,26'NN
				abc="M"
				case 11,27 'Lilith
				abc="N"
				case 12,28'AS
				abc="K"
					'caché si mondial 
					'if choix=4 or choix=5 then carte(i).visible=false
				case 13,29 'FC
				carte(i).visible=false
				abc=""
				case 14,30 'DS
				carte(i).visible=false
				abc=""
				case 15,31 'MC
				abc="L"
					'caché si mondial
					'if choix=4 or choix=5 then carte(i).visible=false
			end select
		'écriture symbole planète
		carte(i).string=abc
		carte(i).CharFontName="Zodiac S" 'mettre après string !
		carte(i).CharHeight=7
		'arrêt clignotement
		carte(i).textanimationkind=0
		'couleur des planètes
			'thème externe
			if i <= 15 then carte(i).charcolor=coul(int(longitudes(i)/30))
			'thème interne
			if i >15 then carte(i).charcolor=coul(int(natal(i-16)/30))
	
	next i


'cercles pour numéros maisons (Sections)
	'effacement des maisons (remove ne marche pas)
	for i= 32 to 43
		carte(i).visible=false
	next i
	
	'pas de maisons si checkbox non coché
	if maisonsok=0 then goto signes

	Point.x = 1500
	Point.y = 500
	Size.Width = 9000
	Size.Height =9000
	
	'sections correspondant aux maisons 1 à 12			
	for i=32 to 43 '44 to 55
		longitude0=maison(i-31)
			if i < 43 then
			longitude1=maison(i-30)
			else
			longitude1=maison(1)
			endif
		'section maison
		carte(i).Position = point
		carte(i).Size = size
			car1=(longitude0+offset)*100 : if car1>36000 then car1=car1-36000
		carte(i).circlestartangle=car1
			car2=car1 + (longitude1-longitude0)*100 : if car2>36000 then car2=car2-36000
		carte(i).circleendangle=car2
			carte(i).CircleKind =  com.sun.star.drawing.CircleKind.SECTION
			'carte(i).CircleKind =  com.sun.star.drawing.CircleKind.CUT
		carte(i).filltransparence=100
		carte(i).linecolor=RGB(220,220,220) 'gris
		'écriture section
		carte(i).visible=true 'mis sinon signe belier non affiché en mondial !
		Page.add(carte(i))	
		'écriture numéro maison	
			'pas de clignotement
			carte(i).textanimationkind=0 
		carte(i).string=str(i-31)
		carte(i).CharHeight=6
	next i



	
signes:	
'cercles pour symboles signes (Cut)
	Point.x = 3000
	Point.y = 2000
	Size.Width = 6000
	Size.Height = 6000
'pour faire des portions de cercle (ex 0-9000,9000-18000,18000-27000 ou 27000-36000 = 1/4 de cercle)sens ccw, 0=axe des x vers la droite
	'début du Bélier
	car1=offset*100
	for i=44 to 55 '32 to 43
		carte(i).Position = point
		carte(i).Size = size
		carte(i).circlestartangle=car1
		car2=car1+3000 : if car2>36000 then car2=car2-36000
		carte(i).circleendangle=car2
			'carte(i).CircleKind =  com.sun.star.drawing.CircleKind.ARC
			'carte(i).CircleKind =  com.sun.star.drawing.CircleKind.SECTION
		carte(i).CircleKind =  com.sun.star.drawing.CircleKind.CUT
			'carte(i).CircleKind =  com.sun.star.drawing.CircleKind.FULL
		'blanc
		carte(i).FillColor = RGB(255, 255, 255)
		carte(i).linecolor=RGB(0,0,0) 'gris
		'carte(i).filltransparence=100
	'	carte(i).linedash.style=6 'pas d'action
		'écriture arc
		carte(i).visible=true
		Page.add(carte(i))
		'écriture symboles signes
		abc=chr$(97+i-44) 'a à l
		carte(i).string=abc
		carte(i).CharFontName="Zodiac S" 'mettre après string !
		carte(i).CharHeight=8
		'couleur du signe
		carte(i).charcolor=coul(i-44)
	car1=car1+3000 : if car1>36000 then car1=car1-36000
	next i
	
		


'************************tracé lignes aspects********************************
' choix=1,2,4 - cercle externe (natal, progressé ou mondial) = longitudes (0 à 15) avec carte (0 à 15)
' choix=3;5 - cercle interne (natal) = natal(0 à 15) avec carte (16 à 31) + maison(1 à 12) avec carte(32 à 43)

	'effacement des lignes aspects sinon superposition d'anciens aspects
	for i= 0 to 50
		Page.remove(ligne(i))
	next i

'indice ligne aspect (0 à 50 max)
num=0
'nombre d'aspects
 car2=0


' définition des limites basses et hautes des planètes + maisons sujets et agents
	'planètes sujet : limites basses et hautes par défaut
		su_deb=0 : su_fin=11
	'planètes agent : limite basse par défaut (ag_deb=0 permet de faire clignoter la conjonction Lune.Soleil (Soleil/Lune ne clignote pas !)
		ag_deb=0
	'planètes agent : limite haute
	select case choix
	'natal
	case 1
		ag_fin=11'15
	'progressé/progressé
	case 2
		ag_fin=11'15 '43
	'progressé/natal
	case 3
		ag_fin=43
	'mondial
	case 4
		'case mondial cochée : thème mondial Mars-Pluton
		if mondialok=1 then
			su_deb=4 : su_fin=9
			ag_deb=5 : ag_fin=9
		'case mondial décochée : toutes planètes
		else
			ag_fin=11'15
		endif
	'mondial/natal
	case 5
		ag_fin=31
	end select
	
'planète sujet	
for m =su_deb to su_fin '0 to 11'15

	'comparaison d'une planète uniquement à celles supérieures (natal, progressé/progressé et mondial) sinon aspects en double
		if choix=1 or choix=2 or choix=4 then ag_deb=m+1
	'pas d'aspects aux axes pour natal (sinon trop d'aspects) et mondial
		if m >11 then if choix=1 or choix=4 or choix=5 then goto fin_m
	'pas d'aspects lune mondiale/ planete natale en mondial/natal
		if choix=5 and m=1 then goto fin_m
	
	
	'planète agent (0-15 : planètes externes, 16-31 : planètes internes, 32-43 : maisons)
	for j=ag_deb to ag_fin
		if j > ag_fin then goto fin_m
			'pas de comparaison d'une planète à elle-même sauf en progressé/natal (choix=3) et mondial/natal
		'	if j=m and choix <>3 and choix <>5 then goto fin_j
			'pas de comparaison aux maisons sauf Soleil et Lune en progressé/natal (j>15, m=0 ou 1 et choix=3)
		'	if j>15 then if m >1 or choix=1 or choix=2 or choix=4 then exit for
			
			'longitude sujet (externe)
				longitude1=longitudes(m)
			'longitude agent (externe ou interne)
				select case choix
				'natal, progressé/progressé, mondial
				case 1,2,4
					longitude0=longitudes(j)
				'progressé/natal ou mondial/natal
				case 3,5
					'planètes externes : pas de comparaison
					if j <=15 then
						goto fin_j
					'planètes internes
					elseif j >15 and j <=31 then
						longitude0=natal(j-16)
					elseif j >31 then
					'maisons internes
						longitude0=maison(j-31)
					endif
				end select
					
			'calcul différence
			gap=longitude0-longitude1
				if gap < 0 then gap=gap+360
			'division par 15 pour une approximation de l'aspect
			val1=int(gap/15)
			
			'pas d'autres aspects que la conjonction (val1=0) à la Lune (m=1) pour les maisons (j>31)
			if j >31 and m >1 and val1 >0 then goto fin_j
				
		'aspects proches de cette approximation			
		for k=arc(val1,0) to arc(val1,1)
				'orbe à comparer à l'orbe maximum de la transitée
				orbedecimal=abs(gap-angle(k))
			
					'déinition de l'orbe
					select case choix
					'natal ou mondial
					case 1,4
						'orbes définis dans orbe_theme(,,) de Soleil à Pluton
						if m < 12 and j < 12 then
							orbemax=orbe_theme(m,j,k mod 16)
						'orbes de la feuille éphémérides de AS à MC
						else
							orbemax = orbe_as_fc(k mod 16)(0)
						endif
					'progressé/progressé (on utilise la tolérance la + grande, utile sinon différence lune-mercure ou mercure-lune par ex.)
					case 2
						orbemax= tol1*oFA.callFunction("Max",array(tolerance(j),tolerance(m)))
					'progressé/natal
					case 3
						orbemax=tol1*tolerance(m)
					'transits mondial/natal
					case 5
						orbemax=tol2*tolerance(m)
					end select	
								
			 '******aspect trouvé*******
			 if orbedecimal <= orbemax then
			 
			 	'écriture aspect (empêche le tracé des aspects !)
		'	if tol2=1/3 then
			' Sheet.getCellByPosition(10,k).string=aspect(k) : Sheet.getCellByPosition(11,k).string=planete(m) : Sheet.getCellByPosition(12,k).string=planete(j)
		'	endif
				
				'foutoir ! clignotement planètes si conjonction (avant si val1=0, arrière si val1=23), en progressé (choix=2 ou 3) et mondial/natal (choix=5)
				if val1=0 or val1=23 then if choix=2 or choix=3 or choix=5 then
					'pas de clignotement entre axes identiques (progressé/natal)
				'	if j=m and j >11 then goto suite
				'	if j >11 or m >11 then goto suite
				'thème externe (sujet)
				 carte(m).textanimationkind=1 
				'thème externe (choix=2) ou interne (choix=3 ou 5) (agent)
						'if choix=2 then carte(j).textanimationkind=1 else carte(j+16).textanimationkind=1
			'	carte(j).textanimationkind=1 'pb avec clignotement maison
				endif
	suite:	
				'pas de tracés pour axes et maisons	(j ou m > 31)
				if m <12 then if j <=31 then
				 car2=car2+1
					astre=array(m,j)
					'calcul des coordonnées à l'écran des 2 planètes en aspect
					for i=0 to 1
						'thème externe
							if i=0 or choix =1 or choix =2 or choix=4 then longitude=longitudes(astre(i))+offset
						'thème interne
							if i=1 and (choix=3 or choix=5) then longitude=natal(astre(i)-16)+offset 'j est forcément >15
								if longitude <0 then longitude=longitude+360
								if longitude >360 then longitude=longitude-360
						select case longitude
							case 0 to 90
								call calc_sinus(choix,i,longitude)
								hauteur(i)=3000-(hauteur(i)/2)
								longueur(i)=3000+longueur(i)/2
							case 90 to 180
								call calc_sinus(choix,i,180-longitude)
								hauteur(i)=3000-(hauteur(i)/2)
								longueur(i)=3000-(longueur(i)/2)
							case 180 to 270
								call calc_sinus(choix,i,longitude-180)
								hauteur(i)=3000+hauteur(i)/2
								longueur(i)=3000-(longueur(i)/2)
							case 270 to 360
								call calc_sinus(choix,i,360-longitude)
								hauteur(i)=3000+hauteur(i)/2
								longueur(i)=3000+longueur(i)/2
						end select
					next i
						
					'tracé ligne aspect
					Point.x=3000+longueur(0)
					Point.y=2000+hauteur(0)
					Size.Width = longueur(1)-longueur(0)
					Size.Height =hauteur(1)-hauteur(0)
					ligne(num).Position = point
					ligne(num).Size = size
					ligne(num).linecolor=couleur(k mod 16)
					Page.add(ligne(num))
					
				num=num+1
				exit for
				endif
			endif	 
		next k
fin_j:
	next j
fin_m:
next m

'affiche nombre d'aspects
sheet.getcellbyposition(0,1).string= "aspects : " & car2

'*******************************************fin aspects********************

'mise au 1er plan de la feuille
	Sheet = Doc.Sheets.getByName("zodiaque")
	Controller = Doc.CurrentController
	controller.setActiveSheet(sheet)
	
'focus sur cellule J1
	cell=Sheet.getCellByPosition(9,0)
	controller.select(cell)

'trop de messages !
'if message <> "" then msgbox message


'********************************attente mise à jour interactive************************
	if form_ok=1 then
	on error goto fin
			do
			
				'*********clic sur date
				if commande1.date.day <> jour or commande1.date.month <> mois or commande1.date.year <> annee then
					'active bouton 'variable' (1=aujourd'hui, 2=anniversaire, 3=variable)
					feuille.getControl(options2(3)).state=1 : option2_indice=3 
					'reset tolérances 1 jour
					if (commande1.date.day <> jour or  commande1.date.month <> mois) then 'and (choix =2 or choix=3)
						feuille.getControl(options3(2)).state=1 : option3_indice=2
						tol1=tol1min : tol2=tol2min
					endif
					'reset de jour/mois/an
					jour=commande1.date.day
					mois=commande1.date.month
					annee=commande1.date.year
					'calcul de la valeur numérique du jour
					abc=str(commande1.date.day) & "/" &  str(commande1.date.month) & "/" & str(commande1.date.year)
					abc=Replace$(abc," ", "")
					date_jour=datevalue(abc)
					'string de la date du jour
					date_string=abc
					'pour le progressé, calcule coefficients jours et annee 
						'valeur numérique du jour anniversaire de l'année
						abc=Replace$(anniversaire & str(annee)," ", "")
						datex=datevalue(abc)
						'progresse : recalcule jours et annee_progresse
						call calc_progresse(datex,date_jour,annee,anniversaire)
					'pour le mondial, ligne éphémérides 
					rang_mondial=date_jour-an1+17
					
				exit do
				endif
			
									
				
				'**************clic sur heure
				if commande2.time.hours <> hh or commande2.time.minutes <> mm then
					'active bouton 'variable' (1=aujourd'hui, 2=anniversaire, 3=variable)
					feuille.getControl(options2(3)).state=1 : option2_indice=3 
					'reset paramètres
					hh=commande2.time.hours
					mm=commande2.time.minutes
					'calcul du coefficient heure (0 à 1)
					abc=str(commande2.time.hours) & ":" &  str(commande2.time.minutes)
					abc=Replace$(abc," ", "")
					heure=timevalue(abc)
					'heure en string
					heure_string= str(hh) & "h " & str(mm) & "'"
					'reset tolérances 1 jour
					feuille.getControl(options3(2)).state=1 : option3_indice=2
					tol1=tol1min : tol2=tol2min
				exit do
				endif
				
				
				
				'**********************clic sur heure de naissance
				if commande3.time.hours <> hh_naissance or commande3.time.minutes <> mm_naissance then
					'reset paramètres
					hh_naissance=commande3.time.hours
					mm_naissance=commande3.time.minutes
					'calcul du coefficient heure (0 à 1)
					abc=str(commande3.time.hours) & ":" &  str(commande3.time.minutes)
					abc=Replace$(abc," ", "")
					heure_naissance=timevalue(abc)
					'heure en string
					heure_naissance_string= str(hh_naissance) & "h " & str(mm_naissance) & "'"
				exit do
				endif
				
				
										
					
				'****************clic checkbox Maisons
				if feuille.getControl("CheckBox1").state <> maisonsok then
					maisonsok=feuille.getControl("CheckBox1").state
				exit do
				endif
				
				
						
				'****************clic checkbox mondial
				if feuille.getControl("CheckBox2").state <> mondialok then
					mondialok=feuille.getControl("CheckBox2").state
				exit do
				endif
				
				
				
				'****************clic boutons radio tolérance
				if feuille.getControl(options3(option3_indice)).state=0 and choix <> 1 and choix <> 4 then
					'récupère le numéro de l'option
					for i=1 to 2
						if feuille.getControl(options3(i)).state<>0 then option3_indice=i : exit for
					next i
					select case option3_indice
					case 1
						'tol1pour progressé (1=1an), tol2 pour mondial/natal (1=6jours)
						tol1=1 : tol2=1
					case 2
						'tol1=60 si 180 ou 380 pas d'aspects affichés !
						tol1=tol1min : tol2=tol2min
					end select
				exit do
				endif
			
			
			
				'****************clic boutons radio aujourd'hui et anniversaire
				if	feuille.getControl(options2(option2_indice)).state=0 then 
				'récupère le numéro de l'option
					for i=1 to 3
						if feuille.getControl(options2(i)).state<>0 then option2_indice=i : exit for
					next i
					select case option2_indice
					'aujourd'hui
					case 1
						'reset paramètres date
							jour=day(now)
							mois=month(now)
							annee=year(now)
						'change l'affichage de la date
							x2.Year = annee
							x2.Month = mois
							x2.Day =jour		
							commande1.Date = x2
						'reset heure
							y = Time()
							y2.Hours = Hour(y)
							y2.Minutes = Minute(y)
							commande2.Time = y2
							hh=Hour(y)
							mm=Minute(y)
						'date du jour : format numérique
							date_jour=datevalue(now)
						'date du jour : format string
							date_string=date(now)
						'heure actuelle
							abc=time(now)
							'coefficient de 0 à 1 (midi=0,5, minuit=1)
							heure=timevalue(abc) 
							'format string	
							heure_string=mid$(abc,1,5)
						'mondial : ligne éphémérides du jour
							rang_mondial=date_jour-an1+17
						'progresse : recalcule jours et annee_progresse
							call calc_progresse(datex,date_jour,annee,anniversaire)
					'anniversaire de l'année
					case 2
						'reset paramètres date
							jour=day(date_anniv)
							mois=month(date_anniv)
							'annee=year(date_anniv)
						'change l'affichage de la date
							x2.Year = annee
							x2.Month = mois
							x2.Day =jour		
							commande1.Date = x2
						'date anniversaire : format numérique
							abc=Replace$(anniversaire  & str(annee)," ", "")
							date_jour=datevalue(abc)
							datex=date_jour
						' date anniversaire : format string
							date_string=abc
						'mondial : ligne éphémérides du jour
							rang_mondial=date_jour-an1+17
						'progresse : recalcule jours et annee_progresse
							call calc_progresse(datex,date_jour,annee,anniversaire)	
					'variable : pas d'action, le bouton est caché
					'case 3
						
					end select
				exit do
				endif
				
				
				'****************clic optionbuttons 1,2,3,4,5
			
				if	feuille.getControl(options(option_indice)).state=0 then
					'récupère le numéro de choix
					for i=1 to 5
						if feuille.getControl(options(i)).state<>0 then option_indice=i : exit for
					next i
					choix=option_indice
					'si on reste en mondial (choix=2-3) ou progressé (choix=4-5) on ne change rien !
					select case choix
						'natal
						case 1
							'désactive date,heure,heure de naissance,checkbox mondial, 3 boutons "maintenant, anniversaire et variable" et 2 boutons tolérance
							call actions_boutons (false,false,false,false,false,false)		
						'progressé
						case 2,3
							'date,heure,heure de naissance,checkbox mondial, 3 boutons "maintenant, anniversaire et variable" et 2 boutons tolérance
							call actions_boutons (true,false,true,false,true,true)
							
							'mondial précédemment
							if choixref >3 then
								'recalcule jours et annee_progresse
								call calc_progresse(datex,date_jour,annee,anniversaire)
							endif								
						'progressé/progressé
						case 4
							'date,heure,heure de naissance,checkbox mondial, 3 boutons "maintenant, anniversaire et variable" et 2 boutons tolérance
							call actions_boutons (true,true,false,true,true,false)														
						'progressé/natal
						case 5
							'date,heure,heure de naissance,checkbox mondial, 3 boutons "maintenant, anniversaire et variable" et 2 boutons tolérance
							call actions_boutons (true,true,true,false,true,true)			
						end select
												
					'reset valeurs de référence
					curseur2ref=curseur2	
					choixref=choix
			exit do
			endif
			
			'**********************clic bouton sortie
			if clic_sortie=1 then goto fin
			
			'attente 1/2 sec.
			wait 500
			loop
		goto debut
	endif
	
fin:
	if form_ok=1 then 
		'cache checkbox, date, heure, etc.
		feuille.getControl("CheckBox1").visible=false
		call actions_boutons (false,false,false,false,false,false)	
		'active boutons tableaux, thème, aspects, etc.
		call actions_boutons2 (true)	
		'cache bouton sortie
		feuille.getControl("sortie").enable=false
	endif
	
end sub


sub actions_boutons (c1,c2,c3,c4,c5,c6)
	if form_ok=1 then
		'date, heure et heure de naissance
		feuille.getControl("DateField1").visible=c1
		feuille.getControl("TimeField1").visible=c2
		feuille.getControl("TimeField2").visible=c3
		'checkbox mondial
		feuille.getControl("CheckBox2").visible=c4
		'boutons radio maintenant, anniversaire et variable (toujours caché)
		feuille.getControl("Bouton1").visible=c5
		feuille.getControl("Bouton2").visible=c5
		feuille.getControl("Bouton3").visible=false
		'boutons radio tolérance
		feuille.getControl("tol1").visible=c6
		feuille.getControl("tol2").visible=c6
	endif
end sub

sub actions_boutons2 (c1)
	if form_ok=1 then
		feuille.getControl("tableau1").enable=c1
		feuille.getControl("tableau2").enable=c1
		feuille.getControl("tableau3").enable=c1
		feuille.getControl("bouton_ephe").enable=c1
		feuille.getControl("bouton_phases").enable=c1
		feuille.getControl("bouton_theme").enable=c1
		feuille.getControl("bouton_aspects").enable=c1
		feuille.getControl("bouton_graphe").enable=c1
		feuille.getControl("bouton_zodiaque").enable=c1
	endif
end sub

sub calc_sinus(choix,i,angle)
dim rayon
dim sinus
dim coeff as double
oFA = createUnoService( "com.sun.star.sheet.FunctionAccess" )
'les valeurs à l'intérieur du cercle sont doubles de celles à l'écran !
' (ex. la position du point à 90deg. par rapport à l'axe des X est de 6000 = diamètre du cercle au lieu du rayon)
'thème externe
if i=0 or (choix <> 3 and choix <>5) then rayon=6000
'thème interne 
if i=1 and choix=3 then rayon=4000 'progressé/natal
if i=1 and choix=5 then rayon=5000 '/transits mondial/natal

'coefficient pour éloigner les symboles du cercle
coeff=1.05

	'carré du rayon 
	d2=oFA.callFunction("Power",array(rayon,2))
	'conversion en radians
	car=oFA.callFunction("Radians",array(angle))
	'sinus de l'angle
	sinus=oFA.callFunction("Sin",array(car))
	'Y
	hauteur(i)=rayon*sinus
	'a2+b2=c2
	h2= oFA.callFunction("Power",array(hauteur(i),2))
	l2=d2-h2
	'X
	longueur(i)=oFA.callFunction("Sqrt",array(l2))
	'application du coefficient
'	hauteur(i)=(coeff+i/30)*hauteur(i)
'	longueur(i)=(coeff+i/30)*longueur(i)
end sub


sub calc_progresse(datex,date_jour,annee,anniversaire)
	'si l'anniversaire de l'année en cours est postérieur, recul d'une année à l'anniverssaire précédent
	if datex > date_jour then annee_progresse=annee-1 else annee_progresse=annee
	 datex=datevalue(Replace$(anniversaire  & str(annee_progresse)," ", ""))
	'nombre de jours entre date anniversaire prédédent et aujourd'hui
	jours=date_jour-datex
end sub


Function charge_feuille(Libname as String, DialogName as String, Optional oLibContainer)
Dim oLib as Object
Dim oLibDialog as Object
Dim oRuntimeDialog as Object
	If IsMissing(oLibContainer ) then
		oLibContainer = DialogLibraries
	End If
	oLibContainer.LoadLibrary(LibName)
	oLib = oLibContainer.GetByName(Libname)
	oLibDialog = oLib.GetByName(DialogName)
	oRuntimeDialog = CreateUnoDialog(oLibDialog)
	charge_feuille() = oRuntimeDialog
End Function

'pour fermer par exemple avec un bouton si le type du bouton est 'standard' (pas OK) et que l'event du bouton a été assigné à cete macro
Sub ExitDialog1
'form_ok=0
'feuille.endExecute()

	clic_sortie=1
'end
End Sub

'sortie programme
Sub ExitDialog2
	form_ok=0
	feuille.endExecute()
	end
End Sub

Sub visuel
	Doc=thiscomponent
	'vérifie présence feuille éphémérides, sinon création
	If not Doc.Sheets.hasByName("éphémérides") Then 
		Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet") 
		Doc.Sheets.insertByName("éphémérides", Sheet)
	endif

'chargement form
feuille= charge_feuille("Standard","Dialog1")

'affichage valeurs actuelles
	'cache bouton sortie
	feuille.getControl("sortie").enable=false
	'plage éphémérides
	Sheet = Doc.Sheets.getByName("éphémérides")
	commande1 = feuille.getControl("Label1")
	commande1.text=sheet.getcellbyposition(14,17).getstring
	'phases ? (oui ou non)
	feuille.getControl("Label12").text=sheet.getcellbyposition(16,17).getstring
	'nom thème
	commande1 = feuille.getControl("Label2")
	commande1.text=sheet.getcellbyposition(19,0).getstring
	'date de naissance
	commande1 = feuille.getControl("Label11")
	commande1.text=Sheet.getCellByPosition(22, 1).getstring
	'transits
	If Doc.Sheets.hasByName("transits2") then
		Sheet = Doc.Sheets.getByName("transits2")
		commande1 = feuille.getControl("Label4")
		commande1.text=sheet.getcellbyposition(7,0).getstring & "  " & sheet.getcellbyposition(8,0).getstring
	endif
	If Doc.Sheets.hasByName("progressé") then
		Sheet = Doc.Sheets.getByName("progressé")
		commande1 = feuille.getControl("Label5")
		commande1.text=sheet.getcellbyposition(7,0).getstring & "  " & sheet.getcellbyposition(8,0).getstring
	endif
	If Doc.Sheets.hasByName("transits1") then
		Sheet = Doc.Sheets.getByName("transits1")
		commande1 = feuille.getControl("Label9")
		commande1.text=sheet.getcellbyposition(7,0).getstring & "  " & sheet.getcellbyposition(8,0).getstring
	endif
	'transits par planètes
	If Doc.Sheets.hasByName("Uranus") then
		Sheet = Doc.Sheets.getByName("Uranus")
		commande1 = feuille.getControl("Label7")
		'suppression du line feed
		abc=sheet.getcellbyposition(0,0).getstring
		abc= Replace$(abc,chr$(10), " ")
		commande1.text=abc
	endif
	If Doc.Sheets.hasByName("Soleil") then
		Sheet = Doc.Sheets.getByName("Soleil")
		commande1 = feuille.getControl("Label8")
		'suppression du line feed
		abc=sheet.getcellbyposition(0,0).getstring
		abc= Replace$(abc,chr$(10), " ")
		commande1.text=abc
	endif
		If Doc.Sheets.hasByName("Mercure") then
		Sheet = Doc.Sheets.getByName("Mercure")
		if instr(1,sheet.getcellbyposition(0,0).getstring,"progressé")=0 then
		commande1 = feuille.getControl("Label10")
		'suppression du line feed
		abc=sheet.getcellbyposition(0,0).getstring
		abc= Replace$(abc,chr$(10), " ")
		commande1.text=abc
		endif
	endif
'affichage form
form_ok=1
feuille.execute()
call zodiaque(0,1)
End Sub

Sub FileButtonSelected
Dim oFileDlg 'File selection dialog
Dim oSettings 'Settings object to get the path settings
Dim sFile As String 'File URL as a string
Dim oFileAccess 'Simple file access object
Dim oFiles 'The File dialog returns an array of selected files

REM Start with the value in the text field
'	commande1 = feuille2.getControl("FileTextField")
'sFile = commande1.Text
sFile="/home/paul/Documents/Astrologie/Ephemerides_Transits"
	If sFile <> "" Then
	sFile = ConvertToURL(sFile)
	Else
	REM The text field is blank, so obtain the path settings
	oSettings = CreateUnoService("com.sun.star.util.PathSettings")
	sFile = oSettings.Work
	End If
REM Dialog to select a file
oFileDlg = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
REM Set the supported filters
oFileDlg.AppendFilter( "All files (*.*)", "*.*" )
oFileDlg.AppendFilter( "TXT files (*.txt)", "*.txt" )
oFileDlg.SetCurrentFilter( "TXT files (*.txt)", "*.txt")
REM Determine if the "file" is a directory or a file!
oFileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oFileAccess.exists(sFile) Then
	REM I do not force the "display directory" to be a folder.
	REM I could do something fancy with this such as
	REM If NOT oFileAccess.isFolder(sFile) Then extract the folder...
	REM but I won't!
	oFileDlg.setDisplayDirectory(sFile)
	End If
REM Execute the File dialog.
	If oFileDlg.execute() Then
	oFiles = oFileDlg.getFiles()
		If UBound(oFiles) >= 0 Then
		sFile = ConvertFromURL(oFiles(0))
		abc=sFile
		'commande1.Text = sFile
		End If
	End If
End Sub


