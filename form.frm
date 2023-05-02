Attribute VB_Name = "frmPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBHelper As New DBHelper
Dim Pedidos As New M6_Pedidos
Dim Empresa As New EmpresaCOM
Dim CuentasCorrientes As New M0_CuentasCorrientes
Dim SubCuentasCorrientes As New M0_SubCuentasCorrientes
Dim LugaresDeRecepcion As New M0_CtasCtesLugDeRecepcion
Dim Depositos As New M0_Depositos
Dim TiposDeMovimientos As New M6_TiposDeMovimientos
Dim Monedas As New M0_Monedas
Dim Leyendas As New M0_Leyendas
Dim CondicionesComerciales As New M6_CondicionesComerciales
Dim Destinos As New M0_Destinos
Dim Fletes As New M4_FletesBUS
Dim TarifasFletes As New M4_TarifasFletes
Dim Transportistas As New M4_Transportistas
Dim Choferes As New M4_Choferes
Dim Camiones As New M4_Camiones
Dim Listas As New M0_Listas
Dim Comisionistas As New M0_Comisionistas
Dim Cobradores As New M0_Cobradores
Dim Articulos As New M6_Articulos
Dim DerivadorFAC As New DerivadorFAC
Dim FormModulo6 As New M6_FormulariosFAC
Dim Formularios As New M0_Formularios
Dim Campanias As New M6_Campanias
Dim CondicionesDeFlete As New M6_CondicionesDeFlete
Dim ListasM6 As New M6_Listas
Dim Mercaderias As New M3_Mercaderias
Dim Operaciones As New M3_Operaciones
Dim Corredores As New M3_Corredores
Dim Puertos As New M3_Puertos
Dim Caratulas As New M3_Caratulas
Dim CaratulasTipos As New M3_CaratulasTipos
Dim UnidadesDeNegocio As New M0_UnidadesDeNegocio
Dim ListaPrecios As New M6_ListaPrecios

Public TipoDeOperacion As EnumTiposDeOperacion
Public TipoDePedido As EnumTiposDePedido

Public IDPedidos As Long, IDFletes As Long, NoEntrar As Boolean

Private Sub Form_Load()

    CentraForm Me
    Screen.MousePointer = vbArrowHourglass
    
    Pedidos.CadenaDeConexion = CadenaDeConexion
    CuentasCorrientes.CadenaDeConexion = CadenaDeConexion
    SubCuentasCorrientes.CadenaDeConexion = CadenaDeConexion
    LugaresDeRecepcion.CadenaDeConexion = CadenaDeConexion
    Depositos.CadenaDeConexion = CadenaDeConexion
    TiposDeMovimientos.CadenaDeConexion = CadenaDeConexion
    Monedas.CadenaDeConexion = CadenaDeConexion
    Leyendas.CadenaDeConexion = CadenaDeConexion
    CondicionesComerciales.CadenaDeConexion = CadenaDeConexion
    Destinos.CadenaDeConexion = CadenaDeConexion
    Fletes.CadenaDeConexion = CadenaDeConexion
    TarifasFletes.CadenaDeConexion = CadenaDeConexion
    Transportistas.CadenaDeConexion = CadenaDeConexion
    Choferes.CadenaDeConexion = CadenaDeConexion
    Camiones.CadenaDeConexion = CadenaDeConexion
    Listas.CadenaDeConexion = CadenaDeConexion
    Comisionistas.CadenaDeConexion = CadenaDeConexion
    Cobradores.CadenaDeConexion = CadenaDeConexion
                Articulos.CadenaDeConexion = CadenaDeConexion
    DerivadorFAC.CadenaDeConexion = CadenaDeConexion
    FormModulo6.CadenaDeConexion = CadenaDeConexion
    Formularios.CadenaDeConexion = CadenaDeConexion
    Campanias.CadenaDeConexion = CadenaDeConexion
    CondicionesDeFlete.CadenaDeConexion = CadenaDeConexion
    ListasM6.CadenaDeConexion = CadenaDeConexion
    Mercaderias.CadenaDeConexion = CadenaDeConexion
    Operaciones.CadenaDeConexion = CadenaDeConexion
    Corredores.CadenaDeConexion = CadenaDeConexion
    Puertos.CadenaDeConexion = CadenaDeConexion
    Caratulas.CadenaDeConexion = CadenaDeConexion
    CaratulasTipos.CadenaDeConexion = CadenaDeConexion
    UnidadesDeNegocio.CadenaDeConexion = CadenaDeConexion
    ListaPrecios.CadenaDeConexion = CadenaDeConexion
    
    AplicaPermisosAFormulario Me, bdcAbm, "34420"
    AplicaPermisosAFormulario Me, bdcCanje, "33"
  
    Me.tabPedidos.Tab = CtrlTab.tabDivision
    Me.tabPedidos.TabState = 1
    
    Me.tabPedidos.TabsPerRow = 5
    
    If TipoDeOperacion = m6Compra Then
        Me.tabPedidos.Tab = CtrlTab.tabcanje
        Me.tabPedidos.TabState = 1
    End If
    
    If TipoDeOperacion = m6Compra And TipoDePedido = m6Ampliacion Then
        Me.tabPedidos.TabsPerRow = 7
    End If
    
    If TipoDeOperacion = m6Venta And TipoDePedido = m6Ampliacion Then
        Me.tabPedidos.TabsPerRow = 7
    End If
    
    ArmaGrillaDivision sprDivisionPedido
    ArmaGrillaDivision sprDivisionRemitos
    
    If TipoDeOperacion = m6Compra Or TipoDePedido = m6Ampliacion Then
        selProveedores.Visible = False
        lblProveedores.Visible = False
        selProveedores.Id = 0
        If TipoDePedido = m6Ampliacion Then
            Me.tabPedidos.Tab = CtrlTab.tabcanje
            Me.tabPedidos.TabState = 1
            Me.tabPedidos.TabsPerRow = Me.tabPedidos.TabsPerRow - 1
        End If
        Me.tabPedidos.Tab = CtrlTab.tabDivision
        Me.tabPedidos.TabState = 1
        Me.tabPedidos.TabsPerRow = Me.tabPedidos.TabsPerRow - 1
    End If
    
    If TipoDeOperacion = m6Venta And TraeParametro("UDM", mbInsumos) = 1 Then
        Me.tabPedidos.Tab = CtrlTab.tabcanje
        If Me.tabPedidos.TabState = 0 Then
            Me.tabPedidos.TabState = 1
            Me.tabPedidos.TabsPerRow = Me.tabPedidos.TabsPerRow - 1
        End If
        Me.tabPedidos.Tab = CtrlTab.tabDivision
        If Me.tabPedidos.TabState = 0 Then
            Me.tabPedidos.TabState = 1
            Me.tabPedidos.TabsPerRow = Me.tabPedidos.TabsPerRow - 1
        End If
    End If
    UsaListaPrecios = TraeParametro("LPP", mbInsumos)
    
    If TipoDeOperacion = m6Compra Or TipoDePedido = m6Ampliacion Then
        sprPedidos.Height = 4695
        HabilitarControles picCompras, False
    End If
    
    Me.selListaPrecios.Visible = False
    Me.lblListaPrecios.Visible = False
    If TipoDeOperacion = m6Venta Then
        If TraeParametro("HLP", mbInsumos) = 1 Then ' Parametro: Habilita uso de Listas de Precios
            If TraeParametro("LPP", mbInsumos) = 1 Then  ' Parametro: Pedidos: Toma Precio de Listas de Precios
                Me.lblListaPrecios.Visible = True
                Me.selListaPrecios.Visible = True
            End If
        End If
    End If
    
    
    If TipoDePedido = m6Original Then
        If TipoDeOperacion = m6Compra Then
            Me.Caption = "Pedido de Compra de Mercader�a"
        ElseIf TipoDeOperacion = m6Venta Then
            Me.Caption = "Pedido de Venta de Mercader�a"
        End If
    Else
        If TipoDeOperacion = m6Compra Then
            Me.Caption = "Ampliaci�n/reducci�n de pedidos de compra"
        ElseIf TipoDeOperacion = m6Venta Then
            Me.Caption = "Ampliaci�n/reducci�n de pedidos de venta"
        End If
        HabilitarControles fraGeneral, False
        HabilitarControles fraFletes, False
        HabilitarControles fraOtros, False
        HabilitarControles fraDetalle, False
        lblCuentas.Enabled = True
        selCuentas.Enabled = True
        lblCantidad.Enabled = True
        txtCantidad.Enabled = True
        Me.Picture1.Visible = False
    End If
    bdcAbm.Estado = "00001"
    
    If TipoDeOperacion = m6Compra Then
        chkGeneraCaratula.Value = ValueTrue
    Else
        chkGeneraCaratula.Value = ValueFalse
    End If
    txtFComprobante.Text = date
    txtFCondicion.Text = date
    txtFVencimiento.Text = date
    txtFComprobanteO.Text = date
    txtFechaContratoDesde.Text = date
    txtFechaContratoHasta.Text = date
    txtFechaContratoAcreditacion.Text = date
    
    With selUnidadesDeNegocio
        .CamposVisibles = "Descripcion"
        .LlenaLista UnidadesDeNegocio.Lista
        .IDDeDefault = Usuarios.IDUnidadesDeNegocio
    End With
    
    With selProveedores
        .CamposVisibles = "Nombre"
        .CampoDelAlias = "Alias"
        '.ClaveEnRegistro = Me.Caption + " - 2 - " + .Name
        .LlenaLista rsCuentasCorrientesEscritura.Clone
        .AnchoDeLista = 4000
        selProveedores_Click
    End With
    
    With selCuentas
        .CamposVisibles = "Nombre"
        .CampoDelAlias = "Alias"
        .ClaveEnRegistro = Me.Caption + " - 1 - " + .Name
        .CantidadDeFilas = 30
        .AnchoDeLista = 4000
        .LlenaLista rsCuentasCorrientesEscritura.Clone
    End With
    
    With selSubCuentas
        .CamposVisibles = "Descripcion"
    End With
    
    With selTiposDePedidos
        .CamposVisibles = "Descripcion"
        .ClaveEnRegistro = Me.Caption + " - 3 - " + .Name
        .CantidadDeFilas = 20
        .LlenaLista ListasM6.ListaTiposDePedidos
    End With
    
    With selLugaresDeRecepcion
        .CamposVisibles = "Descripcion"
    End With
    
    With selCondicionesDeFlete
        .CamposVisibles = "Descripcion"
        .LlenaLista CondicionesDeFlete.Lista
    End With
    selCuentas_LostFocus
    
    With selModalidades
        .CamposVisibles = "Descripcion"
        .LlenaLista Listas.ListaModalidades(mbInsumos, IIf(TipoDeOperacion = m6Compra, -EnumModalidades.m6Directa, 9999))
        .CantidadDeFilas = 30
        selModalidades_Click
    End With
    
    With Me.selCuentasDivision
        .CamposVisibles = "Nombre"
        .CampoDelAlias = "Alias"
        .CantidadDeFilas = 30
        .AnchoDeLista = 4000
        .LlenaLista rsCuentasCorrientesEscritura.Clone
    End With
        
    
    NoEntrar = True
    
    With selArticulos
        .PideAlias = True
        .CamposVisibles = "NombreP"
        .CampoValorInicial = "NombreP"
        .CampoDelAlias = "Alias"
        .CantidadDeFilas = 30
        .AnchoDeLista = 5000
        .LlenaLista rsArticulos.Clone
    End With
    
        
    With selDepositos
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "sin asignar"
        .LlenaLista Depositos.Lista(mbInsumos)
    End With
    
    NoEntrar = False

    With selDestinos
        .CamposVisibles = "Descripcion"
        .LlenaLista Destinos.Lista(mbInsumos)
    End With
    
    Dim Moneda As String
    With selMonedas
        .CamposVisibles = "Descripcion"
        .LlenaLista Monedas.Lista
        Moneda = LeerRegistry("Software\AS\PedidosMoneda", "Moneda")
        If Moneda <> "" Then
            .Id = CInt(Moneda)
            .IDDeDefault = CInt(Moneda)
        Else
            .Id = 1
            .IDDeDefault = 1
        End If
        selMonedas_LostFocus
    End With

    Dim MonedaExpresado As String
    With selExpresadoEn
        .CamposVisibles = "Descripcion"
        .LlenaLista Monedas.Lista
        If selMonedas.Id <> 1 Then
            MonedaExpresado = LeerRegistry("Software\AS\PedidosMonedaExpresado", "MonedaExpresado")
        Else
            MonedaExpresado = 1
        End If
        If MonedaExpresado <> "" Then
            .Id = CInt(MonedaExpresado)
            .IDDeDefault = CInt(MonedaExpresado)
        Else
            .Id = 1
            .IDDeDefault = 1
        End If
        selExpresadoEn_LostFocus
    End With
    
    With selComentarios
        .CamposVisibles = "Descripcion"
        .LlenaLista Leyendas.Lista '(mbInsumos)
    End With
    
    ParametroFCO = CBool(CCur2(TraeParametro("FCO", mbInsumos)))
    
    With selCondicionesComerciales
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"  'NJE 11/12/2017 - TCK 52389: Agrego la opci�n ninguna para el caso de no utilizar este campo
        .ValorInicial = "ninguna"           'NJE 11/12/2017 - TCK 52389: Agrego la opci�n ninguna para el caso de no utilizar este campo
        .LlenaLista CondicionesComerciales.Lista
        .tag = IIf(ParametroFCO And TipoDeOperacion <> m6Compra, "o", "") 'NJE 11/12/2017 - TCK 52389: Es obligatorio o no de acuerdo al par�metro
        If CBool(CCur2(TraeParametro("PCC", mbInsumos))) = True And TipoDeOperacion <> m6Compra Then
            CondicionesComerciales.TomaUnoPorPredeterminado
            .IDDeDefault = CondicionesComerciales.Id
            .Id = CondicionesComerciales.Id
        End If
    End With
    
    With selComisionistas
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "ninguno"
        .LlenaLista Comisionistas.Lista
    End With
    
    With selCobradores
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "ninguno"
        If TipoDeOperacion = m6Venta And TraeParametro("UVC", mbBase) = True Then
            .LlenaLista Comisionistas.Lista
        Else
            .LlenaLista Cobradores.Lista
        End If
    End With
        
    With selTransportistas
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "flete contratado por el " & IIf(TipoDeOperacion = m6Compra, "productor", "proveedor")
        .LlenaLista Transportistas.Lista
    End With
    selTransportistas_Click
    
    With selCampanias
        .CantidadDeFilas = 30
        .CamposVisibles = "Descripcion"
        .IDDeDefault = Campanias.TomaPredeterminada
        .Id = .IDDeDefault
        .LlenaLista Campanias.Lista
    End With
    
    With selMercaderias
        .CamposVisibles = "Nombre"
        .CantidadDeFilas = 100
        .LlenaLista Mercaderias.Lista()
    End With
    
    With selDestinosO
        .CamposVisibles = "Descripcion"
        .LlenaLista Destinos.Lista(mbCereales, 2, EnumTiposDeDestino.m3Canje)
    End With
    
    With selModalidadesO
        .CamposVisibles = "Descripcion"
        .LlenaLista Listas.ListaModalidades(mbCereales)
    End With
    
    With selTiposDeCaratulas
        .CamposVisibles = "Descripcion"
        selTiposDeCaratulas.LlenaLista CaratulasTipos.Lista(IIf(TipoDeOperacion = m6Compra, m3ContratoDeVenta, m3ContratoDeCompra))
    End With
    
    With selCorredores
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "directo"
        .LlenaLista Corredores.Lista
    End With
    
    With selComisionistasO
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "directo"
        .LlenaLista Comisionistas.Lista
    End With
    
    With selPuertos
        .CamposVisibles = "Descripcion"
        .ClaveEnRegistro = .Name + str(Modulo)
        .LlenaLista Puertos.Lista
    End With
    
    txtSucursalO.Text = "0000"
    txtNumeroO.Text = FormatoRG(Pedidos.UltimoNumero(m6Compra) + 1, 8, True, False)
    '    txtNumeroO.Text = FormatoRG(Operaciones.UltimoNumero(m6CompraOrden) + 1, 8, True, False)
    
    ArmaGrillaPedidos sprPedidos
    ArmaGrillaCuerpo sprDetalle
    
    
    ArmaGrillaPedidos sprPedidosC
    ArmaGrillaCuerpo sprDetalleC
    
    txtSucursal.Text = "0001"

    If Formularios.TomaUnoPorUsuario(IIf(TipoDeOperacion = m6Compra, m6PedidoCompra, m6PedidoVenta), Usuarios.Id, m6SinModalidad) = 0 Then
        bdcAbm.Habilitar , , , False
    Else
        bdcAbm.Habilitar False, , , False
    End If
    
    'PDM 26/09/2017 14:51 //T.48123
    If TipoDeOperacion = m6Compra Then
        lblLugarRecepcion.Caption = "Origen"
    ElseIf TipoDeOperacion = m6Venta Then
        lblLugarRecepcion.Caption = "Recepci�n"
    End If
    
    ColoreaObligatorios Me
    
    Screen.MousePointer = vbDefault
        
    If IDPedidos > 0 Then
        Pedidos.TomaUno IDPedidos
        Me.txtSucursal.Text = Pedidos.Sucursal
        Me.txtNumero.Text = Pedidos.Numero
        txtNumero_LostFocus
    End If
    
    If TipoDePedido = m6Original Then
        Me.tabPedidos.Tab = CtrlTab.tabPedidos
        Me.tabPedidos.TabState = 1
    End If

    If TraeParametro("CSD", mbTodosModulos) = True Then
        'PDM 10/04/2019 12:03 //T.16225 Se agrego para que deshabilite el selector de moneda.
        Me.selExpresadoEn.IDDeDefault = selMonedas.Id
        selMonedas.IDDeDefault = 2
        selMonedas.BackColorCombo = mbObligatorio
        selMonedas.Enabled = False
        txtCotizacion.BackColor = mbObligatorio
        selMonedas.tag = "o"
        selMonedas_LostFocus
        
        Me.selExpresadoEn.Enabled = False
        Me.lblExpresadoEn.Enabled = False
        
    End If
    
    If TipoDeOperacion = m6Venta Then
        With Me.selListaPrecios
            .CamposVisibles = "Descripcion"
            .LlenaLista ListaPrecios.Lista
            ListaPrecios.TomaUnoPorPredeterminado
            .IDDeDefault = ListaPrecios.Id
            .Id = ListaPrecios.Id
        End With
    End If
    ParametroWRH = TraeParametro("WRH", mbInsumos)
    If ParametroWRH = True And TipoDeOperacion = m6Venta Then
        With selProveedores
            .CamposVisibles = "Nombre"
            .CampoDelAlias = "Alias"
            .CampoValorInicial = "Nombre"
            .ValorInicial = "empresa"
            .IDDeDefault = 0
            .LlenaLista rsCuentasCorrientesEscritura.Clone
            
            .AnchoDeLista = 4000
        End With
        selProveedores.Enabled = True
        lblProveedores.Caption = "Titular"
        lblProveedores.Enabled = True
        lblProveedores.Visible = True
    End If
    
    
    If (TraeParametro("UFP", mbInsumos) = 1 And TipoDeOperacion = m6Venta) Then 'Or (TraeParametro("UFP", mbInsumos) = 1 And TipoDeOperacion = m6Venta And IDPedidos > 0) Then
        txtFComprobante.Enabled = False
    Else
        txtFComprobante.Enabled = True
    End If

End Sub
