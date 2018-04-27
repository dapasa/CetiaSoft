Attribute VB_Name = "modPantalla"
Public Sub ActualizarPantalla(OpcionActual As String, Formulario As Form, adoTabla As ADODB.Recordset, Optional Extra As String)
    Select Case OpcionActual
        '********************
        Case "Alumnos"
        '********************
            With adoTabla
                'Actualizo el ID
                id_Alumno = !id
                
                'Asigno datos
                Formulario.txtNombre.Text = UCase(!Nombre & "")
                Formulario.txtDireccion.Text = UCase(!Direccion & "")
                Formulario.txtCodPostal.Text = !CodPostal & ""
                Formulario.cboLocalidad.Text = UCase(DEVOLVER_CAMPO(!idLocalidad, adoLocalidades, "Localidades", "Detalle"))
                Formulario.txtTelefono.Text = !Telefono & ""
                Formulario.txtTelefonoLaboral.Text = !TelefonoLaboral & ""
                
                If Left(!Celular, 2) = "15" Then
                    celu_con_guiones = "15-" & Mid(!Celular, 3, 4) & "-" & Mid(!Celular, 7, 4)
                    Formulario.txtCelular.Text = celu_con_guiones
                Else
                    Formulario.txtCelular.Text = !Celular & ""
                End If
                
                Formulario.cboCompaniaCelular.Text = UCase(DEVOLVER_CAMPO(!idCompaniaCelular, adoCompaniasCelular, "CompaniasCelular", "Detalle"))
                Formulario.txtMail.Text = !Mail & ""
                Formulario.cboComoLlego.Text = UCase(DEVOLVER_CAMPO(!idComoLlego, adoComoLlego, "ComoLlego", "Detalle"))
                Formulario.lblFechaAlta.Caption = !FechaAlta & ""
                Formulario.dtpFechaNac = IIf(Not IsNull(!FechaNac), !FechaNac, Date)
                Formulario.cboTipoDoc.Text = DEVOLVER_CAMPO(!idTipoDoc, adoTiposDoc, "TiposDoc", "Detalle")
                Formulario.txtNumDoc.Text = !NumDoc & ""
                Formulario.txtObservaciones.Text = UCase(!Observaciones & "")
                Formulario.cboEmpresa.Text = UCase(DEVOLVER_CAMPO(!idEmpresa, adoEmpresas, "Empresas", "Nombre"))
                
                Formulario.imgFoto.Picture = LoadPicture(App.Path & "\fotos\" & id_Alumno & ".jpg")
            End With
        '********************
        Case "Factura"
        '********************
            With adoTabla
                If EstiloBuscador = "FacturaEmpresas" Then
                    'Actualizo el ID
                    id_Empresa = !id
                    
                    'Asigno datos
                    Formulario.txtNombre.Text = UCase(!Nombre & "")
                    Formulario.txtEmpresa.Text = UCase(!Nombre & "")
                    Formulario.txtCuit.Text = IIf(!Cuit <> "", !Cuit, "")
                    Formulario.cboCondIva.Text = UCase(DEVOLVER_CAMPO(!idCondIva, adoCondIva, "CondIva", "Detalle"))
                    Formulario.txtDireccion.Text = UCase(!Direccion & "")
                ElseIf EstiloBuscador = "FacturaAlumnos" Then
                    'Actualizo el ID
                    id_Alumno = !id
                    
                    'Asigno datos
                    Formulario.txtNombre.Text = UCase(!Nombre & "")
                    Formulario.txtDireccion.Text = UCase(!Direccion & "")
                    Formulario.txtCuit.Text = !NumDoc
                End If
            End With
        '********************
        Case "Cobranza"
        '********************
            With adoTabla
                If EstiloBuscador = "FacturaEmpresas" Then
                    'Actualizo el ID
                    id_Empresa = !id
                    
                    'Asigno datos
                    Formulario.txtEmpresa.Text = UCase(!Nombre & "")
                    Formulario.txtCuit.Text = IIf(!Cuit <> "", !Cuit, "")
                    Formulario.cboCondIva.Text = UCase(DEVOLVER_CAMPO(!idCondIva, adoCondIva, "CondIva", "Detalle"))
                    Formulario.txtDireccion.Text = UCase(!Direccion & "")
                End If
            End With
        '********************
        Case "Empresas"
        '********************
            With adoTabla
                'Actualizo el ID
                id_Empresa = !id
                
                'Asigno datos
                Formulario.txtNombre.Text = UCase(!Nombre & "")
                Formulario.txtCuit.Text = !Cuit & ""
                Formulario.cboCondIva.Text = UCase(DEVOLVER_CAMPO(!idCondIva, adoCondIva, "CondIva", "Detalle"))
                Formulario.txtDireccion.Text = UCase(!Direccion & "")
                Formulario.txtCodPostal.Text = !CodPostal & ""
                Formulario.cboLocalidad.Text = UCase(DEVOLVER_CAMPO(!idLocalidad, adoLocalidades, "Localidades", "Detalle"))
                Formulario.txtTelefono.Text = !Telefono & ""
                Formulario.txtCelular.Text = !Celular & ""
                Formulario.cboCompaniaCelular.Text = UCase(DEVOLVER_CAMPO(!idCompaniaCelular, adoCompaniasCelular, "CompaniasCelular", "Detalle"))
                Formulario.txtMail.Text = !Mail & ""
                Formulario.txtObservaciones.Text = UCase(!Observaciones & "")
            End With
        '********************
        Case "Profesores"
        '********************
            With adoTabla
                'Actualizo el ID
                id_Profesor = !id
                
                'Asigno datos
                Formulario.txtNombre.Text = !Nombre & ""
                Formulario.txtDireccion.Text = !Direccion & ""
                Formulario.txtTelefono.Text = !Telefono & ""
                Formulario.txtCelular.Text = !Celular & ""
                Formulario.cboCompaniaCelular.Text = DEVOLVER_CAMPO(!idCompaniaCelular, adoCompaniasCelular, "CompaniasCelular", "Detalle")
                Formulario.txtMail.Text = !Mail & ""
                Formulario.txtObservaciones.Text = !Observaciones & ""
            End With
        '********************
        Case "Cursos"
        '********************
            With adoTabla
                'Actualizo el ID
                id_Curso = !id
                
                'Asigno datos
                Formulario.lblNumero.Caption = !Numero
                Formulario.cboTipoCurso.Text = !TipoCurso
                Formulario.cboDuracion.Text = !Duracion
                Formulario.optAbierto.Value = !Abierto
                Formulario.optCerrado.Value = Not !Abierto
                Formulario.cboHorario.Text = !Horario
                Formulario.dtpFechaIni.Value = !FechaIni
                Formulario.dtpFechaFin.Value = !FechaFin
                'Formulario.txtValorCuota.Text = !ValorCuota
                Formulario.txtCantCuotas.Text = !CantCuotas
                Formulario.txtVacantes.Text = !Vacantes
                Formulario.lblInscriptos.Caption = !inscriptos
                'Formulario.cboModalidad.Text = DEVOLVER_CAMPO(!idModalidad, adoModalidades, "Modalidades", "Detalle")
                Formulario.cboProfesor.Text = !Profesor
                Formulario.cboAula.Text = !Aula
                Formulario.dtpCuota(1).Value = !Cuota1
                Formulario.dtpCuota(2).Value = !Cuota2
                Formulario.dtpCuota(3).Value = !Cuota3
                Formulario.dtpCuota(4).Value = !Cuota4
                Formulario.dtpCuota(5).Value = !Cuota5
                
                For k = 1 To !CantCuotas
                    If Formulario.dtpCuota(k).Value <> "01/01/2000" Then
                        Formulario.lblCuota(k).Visible = True
                        Formulario.dtpCuota(k).Visible = True
                        Formulario.dtpCuota(k).Tag = Formulario.dtpCuota(k).Value
                    End If
                Next
                
                Formulario.txtObservaciones.Text = !Observaciones & ""
            End With
        '********************
        Case "Inscripcion"
        '********************
            If Extra = "Alumno" Then
                With adoTabla
                    'Actualizo el ID
                    id_Alumno = !id
                    
                    'Asigno datos
                    Formulario.txtAlumno.Text = adoTabla!Nombre
                End With
            End If
    End Select
End Sub
