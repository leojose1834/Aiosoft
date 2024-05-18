Attribute VB_Name = "modulo_registrocompraventa"
'variables utilizadas en el formulario registro de compra
Global DescripcionCodigoProducto(27, 8) As String
Global ValorFila As Integer
Global ValorColumna As Integer
Global TiempoPago As Integer

' variables utilizadas en el formulario detallefactura
Global CodigoDescripcionProducto(27, 16) As String
Global ClienteTiempoPago As Integer
Global LimiteCredito As Double
Global PrecioPorCaja  As Boolean
Global ContadorVenta As Integer

'variable para  los reportes
 Global TipoReporte As Integer
