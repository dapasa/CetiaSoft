VERSION 5.00
Begin VB.Form frmAcumuladoMensual 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmAcumuladoMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CALCULAR_ACUMULADO_MENSUAL()
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE Left(TipoDoc, 2) = 'FC' AND NombreEmisor = '1-Fabru S.A.'"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL FACTURAS FABRU S.A. " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE TipoDoc = 'REC' AND NombreEmisor = '1-Fabru S.A.'"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL FACTURAS FABRU S.A. " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE Left(TipoDoc, 2) = 'FC' AND NombreEmisor = '2-Sergio Fassano'"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL FACTURAS SERGIO FASSANO " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE TipoDoc = 'REC' AND NombreEmisor = '2-Sergio Fassano'"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL FACTURAS SERGIO FASSANO " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE Left(TipoDoc, 2) = 'FC' AND NombreEmisor = '3-Néstror Russaz'"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL FACTURAS NESTOR RUSSAZ " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE TipoDoc = 'REC' AND NombreEmisor = '3-Néstror Russaz'"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL FACTURAS NESTOR RUSSAZ " & Space(5) & adoTabla!t
    adoTabla.Close
    
    MsgBox mensaje
End Sub

Private Sub Form_Load()
    CALCULAR_ACUMULADO_MENSUAL
End Sub
