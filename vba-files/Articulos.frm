VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Articulos 
   Caption         =   "UserForm4"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8535.001
   OleObjectBlob   =   "Articulos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
'ejecutar el procedimiento conecta
Call conecta
Set rscategorias = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rsarticulos = New ADODB.Recordset
'ejecutamos una consulta a la tabla clientes de la base de datos pcventas
  rscategorias.Open "SELECT nombre from Categorias order by Nombre asc;", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  rs2.Open "SELECT count(*) from Categorias", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  rsarticulos.Open "SELECT * from articulos;", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  
  
   cbcategorias.AddItem (rscategorias.Fields("nombre"))
   cbcategorias.Text = rscategorias.Fields("nombre")
    Dim registros As Integer
    registros = Val(rs2.Fields("Expr1000")) - 1
                    
    For x = 1 To registros Step 1
        rscategorias.MoveNext
       cbcategorias.AddItem (rscategorias.Fields("nombre"))
    Next



End Sub
