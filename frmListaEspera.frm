VERSION 5.00
Begin VB.Form frmListaAlumnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTA DE ALUMNOS"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13785
   Icon            =   "frmListaEspera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   13785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInscribir 
      Caption         =   "&Inscribir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   10920
      Picture         =   "frmListaEspera.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Inscribir al curso"
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   12840
      Picture         =   "frmListaEspera.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   11880
      Picture         =   "frmListaEspera.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar"
      Top             =   3240
      Width           =   855
   End
   Begin VB.Frame fraListaAlumnos 
      Caption         =   "Alumnos en espera para el curso "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.ListBox lstListaAlumnos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         ItemData        =   "frmListaEspera.frx":1108
         Left            =   120
         List            =   "frmListaEspera.frx":110A
         TabIndex        =   1
         Top             =   480
         Width           =   13335
      End
      Begin VB.Label lblListaAlumnos 
         Caption         =   "lblListaAlumnos"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   12975
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   13680
      Y1              =   3120
      Y2              =   3120
   End
End
Attribute VB_Name = "frmListaAlumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' MÓDULO: Lista de alumnos            FECHA: Ago / 2007
'******************************************************
' RESUMEN: se utiliza para ver alumnos inscriptos y en
'          espera para un curso determinado.
'******************************************************
' ÚLTIMA MODIFICACIÓN IMPORTANTE: 14/10/2007
'******************************************************
' ETAPA: desarrollo.
'******************************************************
' AUTOR: Pablo Adrián Langholz
' CONTACTO: elmaildepablo@gmail.com
'******************************************************

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lstListaAlumnos.ListIndex <> -1 Then
        If MsgBox("¿Confirma que desea quitar al alumno de la lista de " & TipoLista & "?", vbQuestion + vbYesNo, "QUITAR ALUMNO") = vbYes Then
            id_Alumno = Val(Left(lstListaAlumnos.List(lstListaAlumnos.ListIndex), 6))
            id_Curso = Val(Left(frmInscripcion.lstCursos.List(frmInscripcion.lstCursos.ListIndex), 5))
            
            lstListaAlumnos.RemoveItem (lstListaAlumnos.ListIndex)
            
            If TipoLista = "Inscriptos" Then
                sSql = "DELETE FROM AlumnosXCurso WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
                adoConnection.Execute sSql
                
                sSql = "UPDATE Cursos SET Vacantes = Vacantes + 1, Inscriptos = Inscriptos - 1 WHERE id = " & id_Curso
                adoConnection.Execute sSql
                            
                ELIMINAR_DOCUMENTOS
                
                banVolverAListaAlumnos = True
            ElseIf TipoLista = "Espera" Then
                sSql = "DELETE FROM ListaEspera WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
                adoConnection.Execute sSql
            End If
        End If
    Else
        MsgBox "Debe seleccionar un alumno.", vbExclamation, "LISTA DE ALUMNOS"
    End If
End Sub

Private Sub lstListaAlumnos_Click()
    If lstListaAlumnos.ListIndex <> -1 Then
        cmdEliminar.Enabled = True
    Else
        cmdEliminar.Enabled = False
    End If
End Sub

Private Sub ELIMINAR_DOCUMENTOS()
    With adoMovimientos
        CERRAR_TABLA adoMovimientos
        sSql = "SELECT * FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
        .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                sSql = "DELETE FROM ItemsXMov WHERE idMovimiento = " & !id
                adoConnection.Execute sSql
                
                '.Delete
                
                .MoveNext
            Loop
            
            .Close
        End If
        
        sSql = "DELETE FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
        adoConnection.Execute sSql
    End With
End Sub
