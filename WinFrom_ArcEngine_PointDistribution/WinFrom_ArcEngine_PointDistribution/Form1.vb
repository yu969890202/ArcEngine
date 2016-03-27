Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geometry.esriGeometryType
Imports ESRI.ArcGIS.Carto.esriViewDrawPhase
Imports ESRI.ArcGIS.Controls

Public Class Form1
    Private pPointFeatureLayer As IFeatureLayer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AxMapControl1.LoadMxFile(Application.StartupPath + "\\Map\\Map.mxd")
    End Sub
    Private Sub Bt_OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bt_OK.Click
        pPointFeatureLayer = AxMapControl1.get_Layer(0)
        CreateFeatureClass()
    End Sub

#Region "Create FeatureClass"

    ''' <summary>
    ''' 创建渲染图层
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateFeatureClass() As IFeatureClass
        Dim pWorkspaceFactory As IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesGDB.AccessWorkspaceFactory
        Dim dataset As IDataset = pPointFeatureLayer.FeatureClass
        Dim featureWorkspace As IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(dataset.Workspace.PathName, 0) 'dataset.Workspace.PathName="C:\\aa.mdb"
        Dim workspace As IWorkspace2 = featureWorkspace
        Dim featureClassName As String = "New"
        Dim fields As IFields = Nothing
        Dim CLSID As ESRI.ArcGIS.esriSystem.UID = Nothing
        Dim CLSEXT As ESRI.ArcGIS.esriSystem.UID = Nothing
        Dim strConfigKeyword As String = ""
        Dim featureClass As IFeatureClass
        If workspace.NameExists(esriDatasetType.esriDTFeatureClass, featureClassName) Then '如果存在删除改要素
            featureClass = featureWorkspace.OpenFeatureClass(featureClassName)
            Dim pDataset As IDataset = featureClass
            pDataset.Delete() '删除该要素
        End If
        ' 赋值类ID如果未分配
        If CLSID Is Nothing Then
            CLSID = New ESRI.ArcGIS.esriSystem.UID
            CLSID.Value = "esriGeoDatabase.Feature"
        End If
        Dim objectClassDescription As IObjectClassDescription = New FeatureClassDescription
        If fields Is Nothing Then
            ' 创建字段
            fields = objectClassDescription.RequiredFields
            Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)
            Dim field As IField = New Field
            Dim fieldEdit As IFieldEdit = CType(field, IFieldEdit) ' 显示转换
            ' 设置字段属性
            fieldEdit.Name_2 = "num"
            fieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
            fieldEdit.IsNullable_2 = False
            fieldEdit.AliasName_2 = "num"
            fieldEdit.DefaultValue_2 = 0
            fieldEdit.Editable_2 = True
            '添加到字段集中
            fieldsEdit.AddField(field)
            fields = CType(fieldsEdit, IFields)
        End If
        Dim strShapeField As String = ""
        Dim j As Int32
        For j = 0 To fields.FieldCount
            If fields.Field(j).Type = esriFieldType.esriFieldTypeGeometry Then
                strShapeField = fields.Field(j).Name
                Exit For
            End If
        Next j
        Dim fieldChecker As IFieldChecker = New FieldChecker
        Dim enumFieldError As IEnumFieldError = Nothing
        Dim validatedFields As IFields = Nothing
        fieldChecker.ValidateWorkspace = CType(workspace, IWorkspace)
        fieldChecker.Validate(fields, enumFieldError, validatedFields)
        featureClass = featureWorkspace.CreateFeatureClass(featureClassName, validatedFields, CLSID, CLSEXT, esriFeatureType.esriFTSimple, strShapeField, strConfigKeyword)

        '添加要素，跟据其他要素的范围，生成一个边长为length的矩形网格用于渲染
        Dim length As Integer = 1000000 '方块长度
        Dim pEnvelope As IEnvelope = pPointFeatureLayer.AreaOfInterest
        Dim XMin As Double = pEnvelope.XMin
        Dim XMax As Double = pEnvelope.XMax + length

        Dim YMax As Double = pEnvelope.YMax
        Dim newXMin As Double = XMin + length

        Do While newXMin < XMax
            Dim YMin As Double = pEnvelope.YMin
            Dim newYMin As Double = YMin + length
            Do While newYMin < YMax
                AddFeature(XMin, newXMin, YMin, newYMin, featureClass.CreateFeature())
                YMin = newYMin
                newYMin = YMin + length
            Loop
            AddFeature(XMin, newXMin, YMin, newYMin, featureClass.CreateFeature())
            XMin = newXMin
            newXMin = XMin + length
        Loop
        Dim pNewFeatureLayer As IFeatureLayer = New FeatureLayer
        pNewFeatureLayer.FeatureClass = featureClass
        DefineUniqueValueRenderer(pNewFeatureLayer, "num")
        AxMapControl1.AddLayer(pNewFeatureLayer, 1)
        AxMapControl1.Refresh()
        MessageBox.Show("生成完毕")
    End Function
    '创建方格要素
    Private Function AddFeature(ByVal XMin As Double, ByVal newXMin As Double, ByVal YMin As Double, ByVal newYMin As Double, ByVal pFeature As IFeature)
        Dim pPoint1 As IPoint = New Point()
        pPoint1.X = XMin
        pPoint1.Y = YMin
        Dim pPoint2 As IPoint = New Point()
        pPoint2.X = newXMin
        pPoint2.Y = YMin
        Dim pPoint3 As IPoint = New Point()
        pPoint3.X = newXMin
        pPoint3.Y = newYMin
        Dim pPoint4 As IPoint = New Point()
        pPoint4.X = XMin
        pPoint4.Y = newYMin
        Dim pPolygon As IPolygon
        Dim pPointColec As IPointCollection = New Polygon
        pPointColec.AddPoint(pPoint1)
        pPointColec.AddPoint(pPoint2)
        pPointColec.AddPoint(pPoint3)
        pPointColec.AddPoint(pPoint4)
        pPolygon = CType(pPointColec, IPolygon)
        pFeature.Shape = pPolygon
        Dim pSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
        pSpatialFilter.Geometry = pPolygon
        pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects '相交的状态
        Dim featureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor = pPointFeatureLayer.FeatureClass.Search(pSpatialFilter, False)
        Dim count As Integer = 0
        Dim pTmpFeature = featureCursor.NextFeature()
        While Not IsNothing(pTmpFeature)
            count += 1
            pTmpFeature = featureCursor.NextFeature()
        End While
        Dim fieldindex As Integer = pFeature.Fields.FindField("num")
        pFeature.Value(fieldindex) = count
        pFeature.Store()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(featureCursor)
    End Function

    '唯一值渲染
    Private Sub DefineUniqueValueRenderer(ByVal pGeoFeatureLayer As IGeoFeatureLayer, ByVal fieldName As String)
        '创建于渲染器的符号渐变颜色.
        '创建渐变色带
        Dim algColorRamp As IAlgorithmicColorRamp = New AlgorithmicColorRamp
        algColorRamp.FromColor = getRgbColor(245, 245, 245) '灰色
        algColorRamp.ToColor = getRgbColor(255, 0, 0)
        algColorRamp.Algorithm = esriColorRampAlgorithm.esriCIELabAlgorithm

        Dim pUniqueValueRenderer As IUniqueValueRenderer = New UniqueValueRenderer()

        Dim pSimpleFillSymbol As ISimpleFillSymbol = New SimpleFillSymbol()
        pSimpleFillSymbol.Style = esriSimpleFillStyle.esriSFSSolid
        pSimpleFillSymbol.Outline.Width = 0.4

        '这些属性之前应增加值来设置.
        pUniqueValueRenderer.FieldCount = 1
        pUniqueValueRenderer.Field(0) = fieldName
        pUniqueValueRenderer.DefaultSymbol = pSimpleFillSymbol
        pUniqueValueRenderer.UseDefaultSymbol = True

        Dim pDisplayTable As IDisplayTable = pGeoFeatureLayer
        Dim pFeatureCursor As IFeatureCursor = pDisplayTable.SearchDisplayTable(Nothing, False)

        Dim pFeature As IFeature = pFeatureCursor.NextFeature()
        Dim ValFound As Boolean
        Dim fieldIndex As Integer

        Dim pFields As IFields = pFeatureCursor.Fields
        fieldIndex = pFields.FindField(fieldName)

        While Not pFeature Is Nothing
            Dim pClassSymbol As ISimpleFillSymbol = New SimpleFillSymbol
            pClassSymbol.Style = esriSimpleFillStyle.esriSFSSolid
            pClassSymbol.Outline.Width = 0.4
            Dim classValue As String
            classValue = pFeature.Value(fieldIndex)
            '测试以查看是否该值被添加。如果没有就添加
            ValFound = False
            Dim i As Integer
            For i = 0 To pUniqueValueRenderer.ValueCount - 1 Step i + 1
                If pUniqueValueRenderer.Value(i) = classValue Then
                    ValFound = True
                    Exit For
                End If
            Next
            If ValFound = False Then
                pUniqueValueRenderer.AddValue(classValue, fieldName, pClassSymbol)
                pUniqueValueRenderer.Label(classValue) = classValue
                pUniqueValueRenderer.Symbol(classValue) = pClassSymbol
            End If
            pFeature = pFeatureCursor.NextFeature()
        End While
        algColorRamp.Size = pUniqueValueRenderer.ValueCount
        Dim bOK As Boolean
        algColorRamp.CreateRamp(bOK)
        Dim pEnumColors As IEnumColors = algColorRamp.Colors
        pEnumColors.Reset()
        Dim j As Integer
        For j = 0 To pUniqueValueRenderer.ValueCount - 1 Step j + 1
            Dim xv As String
            xv = pUniqueValueRenderer.Value(j)
            If xv <> "" Then
                Dim pSimpleFillColor As ISimpleFillSymbol = pUniqueValueRenderer.Symbol(xv)
                pSimpleFillColor.Color = pEnumColors.Next()
                pUniqueValueRenderer.Symbol(xv) = pSimpleFillColor
            End If
        Next
        pUniqueValueRenderer.ColorScheme = "Custom"
        Dim pTable As ITable = pDisplayTable
        Dim isString As Boolean = pTable.Fields.Field(fieldIndex).Type = esriFieldType.esriFieldTypeString
        pUniqueValueRenderer.FieldType(0) = isString
        pGeoFeatureLayer.Renderer = pUniqueValueRenderer
        Dim pUID As IUID = New UID()
        pUID.Value = "{683C994E-A17B-11D1-8816-080009EC732A}"
        pGeoFeatureLayer.RendererPropertyPageClassID = pUID
    End Sub

    Private Function getRgbColor(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As IColor
        Dim pRgbColr As IRgbColor = New RgbColor
        pRgbColr.Red = r
        pRgbColr.Green = g
        pRgbColr.Blue = b
        Dim pColor As IColor = CType(pRgbColr, IColor)
        Return pColor
    End Function

#End Region

   
End Class
