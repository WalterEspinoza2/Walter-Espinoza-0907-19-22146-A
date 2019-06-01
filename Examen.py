Private Sub AMARILLO_Click() 
TEXTO.BackColor = vbYellow 
End Sub 
este codigo es para darle colo a un texto 
este otro codigo es para el tamaño de letra 
Private Sub PEQUEÑO_Click() 
TEXTO.FontSize = 8 
End Sub 
este otro codigo es para el formato de texto 



Private Sub CASTELLAR_Click() 
TEXTO.FontName = "CASTELLAR" 
End Sub 
este oro codigo es para suma de n numeros 
Private Sub Form_Load() 
'DECLARACION DE VARIABLES 
Dim INGRESADO As Integer 
Dim NUMERO As Integer 
Dim CONTADOR As Integer 
Dim SUMA As Integer 
SUMA = 0 
INGRESADOS = InputBox("CUANTAS VECES SE VA A REPETIR:") 

For CONTADOR = 1 To INGRESADOS 
NUMERO = InputBox("ESCRIBA UN NUMERO:") 
SUMA = SUMA + NUMERO 

Next CONTADOR 
PROMEDIO = SUMA / INGRESADOS 
TOTAL = SUMA 
LIEDOS = INGRESADOS 
End Sub 


Este es la vincualcion de bisual basic con base de tados botones de accion(microsotf access) 
Private Sub buscar_Click() 
Dim buscar As String, criterio As String 
buscar = InputBox("¿Que nombre desea buscar?", "busqueda por nombre", vbQuestion) 
If buscar = "" Then Exit Sub 
criterio = "nombre like '*" & buscar & "*'" 
Adodc1.Recordset.MoveNext 
If Not Adodc1.Recordset.EOF Then 
Adodc1.Recordset.Find criterio 
End If 
If Adodc1.Recordset.EOF Then 
Adodc1.Recordset.MoveFirst 
Adodc1.Recordset.Find criterio 
If Adodc1.Recordset.EOF Then 
Adodc1.Recordset.MoveLast 
respuesta = MsgBox("Alumno no encontrado", vbCritical) 
End If 
End If 
End Sub 

Private Sub cancelar_Click() 
Adodc1.Recordset.CancelUpdate 
ver.Visible = False 
Calendario.Visible = False 
fecha_nac.Visible = True 
sex.Visible = False 
sexo.Visible = True 
End Sub 


Private Sub eliminar_Click() 
Dim confirmacion As Integer 
confirmacion = MsgBox("¿Desea eliminar el Alumno?", vbYesNo + vbQuestion + vbDefaultButton2, "Eliminar Alumno") 
If confirmacion = vbYes Then 
Adodc1.Recordset.Delete 
MsgBox ("El alumno ha sido borrado") 
Adodc1.Recordset.MoveNext 
If Adodc1.Recordset.EOF Then 
Adodc1.Recordset.MoveLast 
End If 
Else 
Exit Sub 
End If 
End Sub 

Private Sub guardar_Click() 
Adodc1.Recordset.Update 
mensaje = MsgBox("El alumno ha sido guardado exitosamente") 
ver.Visible = False 
Calendario.Visible = False 
fecha_nac.Visible = True 
End Sub 

Private Sub nuevo_Click() 
Adodc1.Recordset.AddNew 
guardar.Visible = True 
cancelar.Visible = True 
sexo.Visible = False 
sex.Visible = True 
ver.Visible = True 
fecha_nac.Visible = False 
End Sub 

Private Sub salir_Click() 
End 
End Sub 

Private Sub ver_Click() 
Calendario.Visible = True 
End Sub 
Private Sub inicio_Click() 
Adodc1.Recordset.MoveFirst 
End Sub 

Private Sub siguiente_Click() 
Adodc1.Recordset.MoveNext 
If Adodc1.Recordset.EOF Then 
Adodc1.Recordset.MoveLast 
End If 
End Sub 
Private Sub final_Click() 
Adodc1.Recordset.MoveLast 
End Sub 
Private Sub anterior_Click() 
Adodc1.Recordset.MovePrevious 
If Adodc1.Recordset.BOF Then 
Adodc1.Recordset.MoveFirst 
End If 
End Sub