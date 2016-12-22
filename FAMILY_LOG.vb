Imports system
Imports system.IO
Imports system.xml
Imports System.Net
Imports System.Windows.Forms
Imports System.ComponentModel
Imports Newtonsoft.Json

Namespace OData
    
    <QueryTitle("Part Families")>  _
    Public Class QUERY_FAMILY_LOG
        Inherits oDataQuery
        
        Private _Name As String
        
        Private _Value As list(of FAMILY_LOG)
        
        Private _Parent As oDataObject
        
        Private _BindingSource As System.Windows.Forms.BindingSource
        
        Public Sub New()
            MyBase.New
            _Value = new list (of FAMILY_LOG)
            _Parent = nothing
            _Name = "FAMILY_LOG"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            .add(0, "Parts in Family")
            .add(1, "Malfunction Codes for Family")
            .add(2, "Service Terms per Family")
            .add(3, "Required Analyses")
            end with
        End Sub
        
        Public Sub New(ByRef Parent As oDataObject)
            MyBase.New
            _Value = new list (of FAMILY_LOG)
            _Parent = Parent
            _name = "FAMILY_LOG_SUBFORM"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            .add(0, "Parts in Family")
            .add(1, "Malfunction Codes for Family")
            .add(2, "Service Terms per Family")
            .add(3, "Required Analyses")
            end with
        End Sub
        
        <JsonProperty("@odata.context")>  _
        Protected Friend Property Name() As String
            Get
                return _Name
            End Get
            Set
                _Name = value
            End Set
        End Property
        
        Public Property Value() As list(of FAMILY_LOG)
            Get
                return _Value
            End Get
            Set
                _Value = value
            End Set
        End Property
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                return _Name
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        Public Overrides ReadOnly Property ObjectType() As System.Type
            Get
                return GetType(FAMILY_LOG)
            End Get
        End Property
        
        Public Overrides ReadOnly Property BindingSource() As System.Windows.Forms.BindingSource
            Get
                Return _BindingSource
            End Get
        End Property
        
        Protected Friend Overrides Sub Deserialise(ByRef Stream As StreamReader)
            _value.clear
            For Each _FAMILY_LOG As FAMILY_LOG In JsonConvert.DeserializeObject(Of QUERY_FAMILY_LOG)(stream.ReadToEnd).Value
              With _FAMILY_LOG
                  .Parent = Parent
                  .loading = false
              end with
              _value.add(_FAMILY_LOG)
            next
            _BindingSource.DataSource = _Value
        End Sub
        
        Protected Friend Overrides Sub HandlesAdd(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILY_LOG = JsonConvert.DeserializeObject(Of FAMILY_LOG)(e.StreamReader.ReadToEnd)
                With TryCast(BindingSource.Current, FAMILY_LOG)
                  .FAMILYNAME = obj.FAMILYNAME
                  .FAMILYDESC = obj.FAMILYDESC
                  .EFAMILYDES = obj.EFAMILYDES
                  .TECHFLAG = obj.TECHFLAG
                  .DEBITFLAG = obj.DEBITFLAG
                  .PROJPROCESS = obj.PROJPROCESS
                  .TECHNICIANLOGIN = obj.TECHNICIANLOGIN
                  .FTCODE = obj.FTCODE
                  .FTNAME = obj.FTNAME
                  .DISTRFLAG = obj.DISTRFLAG
                  .DUTYPERCENT = obj.DUTYPERCENT
                  .DUTYPERCENT1 = obj.DUTYPERCENT1
                  .DUTYPERCENT2 = obj.DUTYPERCENT2
                  .DUTYPERCENTTYPE = obj.DUTYPERCENTTYPE
                  .SELLFLAG = obj.SELLFLAG
                  .NEGBALFLAG = obj.NEGBALFLAG
                  .FORECAST = obj.FORECAST
                  .PARTNAME = obj.PARTNAME
                  .PARTDES = obj.PARTDES
                  .STORAGETYPECODE = obj.STORAGETYPECODE
                  .STORAGETYPEDES = obj.STORAGETYPEDES
                  .SERIALMONTHEXPIRY = obj.SERIALMONTHEXPIRY
                  .ISMFAMILY = obj.ISMFAMILY
                  .MFAMILYNAME = obj.MFAMILYNAME
                  .MFAMILYDES = obj.MFAMILYDES
                  .INACTIVE = obj.INACTIVE
                  .FAMILY = obj.FAMILY
                  .loading = false
                end with
            End If
        End Sub
        
        Private Sub HandlesBeginAdd(ByVal sender As Object, ByVal e As ComponentModel.AddingNewEventArgs)
            e.NewObject = new FAMILY_LOG(me.Parent)
        End Sub
        
        Protected Friend Overrides Sub Remove(ByRef obj As oDataObject)
            for each _FAMILY_LOG as FAMILY_LOG in value
              If _FAMILY_LOG.Equals(trycast(obj,FAMILY_LOG)) Then
                  value.remove(_FAMILY_LOG)
                  exit for
              end if
            next
        End Sub
    End Class
    
    Public Class FAMILY_LOG
        Inherits oDataObject
        
        Private _Parent As oDataObject
        
        Private _IsSetFAMILYNAME As Boolean = Boolean.FalseString
        
        Private _FAMILYNAME As String
        
        Private _IsSetFAMILYDESC As Boolean = Boolean.FalseString
        
        Private _FAMILYDESC As String
        
        Private _IsSetEFAMILYDES As Boolean = Boolean.FalseString
        
        Private _EFAMILYDES As String
        
        Private _IsSetTECHFLAG As Boolean = Boolean.FalseString
        
        Private _TECHFLAG As String
        
        Private _IsSetDEBITFLAG As Boolean = Boolean.FalseString
        
        Private _DEBITFLAG As String
        
        Private _IsSetPROJPROCESS As Boolean = Boolean.FalseString
        
        Private _PROJPROCESS As String
        
        Private _IsSetTECHNICIANLOGIN As Boolean = Boolean.FalseString
        
        Private _TECHNICIANLOGIN As String
        
        Private _IsSetFTCODE As Boolean = Boolean.FalseString
        
        Private _FTCODE As String
        
        Private _FTNAME As String
        
        Private _IsSetDISTRFLAG As Boolean = Boolean.FalseString
        
        Private _DISTRFLAG As String
        
        Private _IsSetDUTYPERCENT As Boolean = Boolean.FalseString
        
        Private _DUTYPERCENT As Decimal
        
        Private _IsSetDUTYPERCENT1 As Boolean = Boolean.FalseString
        
        Private _DUTYPERCENT1 As Decimal
        
        Private _IsSetDUTYPERCENT2 As Boolean = Boolean.FalseString
        
        Private _DUTYPERCENT2 As Decimal
        
        Private _IsSetDUTYPERCENTTYPE As Boolean = Boolean.FalseString
        
        Private _DUTYPERCENTTYPE As String
        
        Private _IsSetSELLFLAG As Boolean = Boolean.FalseString
        
        Private _SELLFLAG As String
        
        Private _IsSetNEGBALFLAG As Boolean = Boolean.FalseString
        
        Private _NEGBALFLAG As String
        
        Private _IsSetFORECAST As Boolean = Boolean.FalseString
        
        Private _FORECAST As String
        
        Private _IsSetPARTNAME As Boolean = Boolean.FalseString
        
        Private _PARTNAME As String
        
        Private _PARTDES As String
        
        Private _IsSetSTORAGETYPECODE As Boolean = Boolean.FalseString
        
        Private _STORAGETYPECODE As String
        
        Private _STORAGETYPEDES As String
        
        Private _IsSetSERIALMONTHEXPIRY As Boolean = Boolean.FalseString
        
        Private _SERIALMONTHEXPIRY As String
        
        Private _IsSetISMFAMILY As Boolean = Boolean.FalseString
        
        Private _ISMFAMILY As String
        
        Private _IsSetMFAMILYNAME As Boolean = Boolean.FalseString
        
        Private _MFAMILYNAME As String
        
        Private _MFAMILYDES As String
        
        Private _IsSetINACTIVE As Boolean = Boolean.FalseString
        
        Private _INACTIVE As String
        
        Private _IsSetFAMILY As Boolean = Boolean.FalseString
        
        Private _FAMILY As Long
        
        Private _FAMILY_LOGPART_SUBFORM As QUERY_FAMILY_LOGPART
        
        Private _FAMILYMALF_SUBFORM As QUERY_FAMILYMALF
        
        Private _FAMILYSERVTYPES_SUBFORM As QUERY_FAMILYSERVTYPES
        
        Private _FAMILYREQLABANALYSES_SUBFORM As QUERY_FAMILYREQLABANALYSES
        
        Public Sub New(ByVal Parent As oDataObject)
            MyBase.New
            _parent = Parent
            ChildQuery.add(0, new oNavigation("Parts in Family"))
            ChildQuery.add(1, new oNavigation("Malfunction Codes for Family"))
            ChildQuery.add(2, new oNavigation("Service Terms per Family"))
            ChildQuery.add(3, new oNavigation("Required Analyses"))
            _FAMILY_LOGPART_SUBFORM = new QUERY_FAMILY_LOGPART(me)
            _FAMILYMALF_SUBFORM = new QUERY_FAMILYMALF(me)
            _FAMILYSERVTYPES_SUBFORM = new QUERY_FAMILYSERVTYPES(me)
            _FAMILYREQLABANALYSES_SUBFORM = new QUERY_FAMILYREQLABANALYSES(me)
            WITH ChildQuery(0)
               .setoDataQuery(_FAMILY_LOGPART_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Parts in Family", _FAMILY_LOGPART_SUBFORM))
                   .add(1, new oNavigation("Malfunction Codes for Family", _FAMILYMALF_SUBFORM))
                   .add(2, new oNavigation("Service Terms per Family", _FAMILYSERVTYPES_SUBFORM))
                   .add(3, new oNavigation("Required Analyses", _FAMILYREQLABANALYSES_SUBFORM))
               end with
            end with
            WITH ChildQuery(1)
               .setoDataQuery(_FAMILYMALF_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Parts in Family", _FAMILY_LOGPART_SUBFORM))
                   .add(1, new oNavigation("Malfunction Codes for Family", _FAMILYMALF_SUBFORM))
                   .add(2, new oNavigation("Service Terms per Family", _FAMILYSERVTYPES_SUBFORM))
                   .add(3, new oNavigation("Required Analyses", _FAMILYREQLABANALYSES_SUBFORM))
               end with
            end with
            WITH ChildQuery(2)
               .setoDataQuery(_FAMILYSERVTYPES_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Parts in Family", _FAMILY_LOGPART_SUBFORM))
                   .add(1, new oNavigation("Malfunction Codes for Family", _FAMILYMALF_SUBFORM))
                   .add(2, new oNavigation("Service Terms per Family", _FAMILYSERVTYPES_SUBFORM))
                   .add(3, new oNavigation("Required Analyses", _FAMILYREQLABANALYSES_SUBFORM))
               end with
            end with
            WITH ChildQuery(3)
               .setoDataQuery(_FAMILYREQLABANALYSES_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Parts in Family", _FAMILY_LOGPART_SUBFORM))
                   .add(1, new oNavigation("Malfunction Codes for Family", _FAMILYMALF_SUBFORM))
                   .add(2, new oNavigation("Service Terms per Family", _FAMILYSERVTYPES_SUBFORM))
                   .add(3, new oNavigation("Required Analyses", _FAMILYREQLABANALYSES_SUBFORM))
               end with
            end with
        End Sub
        
        Public Sub New()
            MyBase.New
            ChildQuery.add(0, new oNavigation("Parts in Family"))
            ChildQuery.add(1, new oNavigation("Malfunction Codes for Family"))
            ChildQuery.add(2, new oNavigation("Service Terms per Family"))
            ChildQuery.add(3, new oNavigation("Required Analyses"))
            _FAMILY_LOGPART_SUBFORM = new QUERY_FAMILY_LOGPART(me)
            _FAMILYMALF_SUBFORM = new QUERY_FAMILYMALF(me)
            _FAMILYSERVTYPES_SUBFORM = new QUERY_FAMILYSERVTYPES(me)
            _FAMILYREQLABANALYSES_SUBFORM = new QUERY_FAMILYREQLABANALYSES(me)
            WITH ChildQuery(0)
               .setoDataQuery(_FAMILY_LOGPART_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Parts in Family", _FAMILY_LOGPART_SUBFORM))
                   .add(1, new oNavigation("Malfunction Codes for Family", _FAMILYMALF_SUBFORM))
                   .add(2, new oNavigation("Service Terms per Family", _FAMILYSERVTYPES_SUBFORM))
                   .add(3, new oNavigation("Required Analyses", _FAMILYREQLABANALYSES_SUBFORM))
               end with
            end with
            WITH ChildQuery(1)
               .setoDataQuery(_FAMILYMALF_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Parts in Family", _FAMILY_LOGPART_SUBFORM))
                   .add(1, new oNavigation("Malfunction Codes for Family", _FAMILYMALF_SUBFORM))
                   .add(2, new oNavigation("Service Terms per Family", _FAMILYSERVTYPES_SUBFORM))
                   .add(3, new oNavigation("Required Analyses", _FAMILYREQLABANALYSES_SUBFORM))
               end with
            end with
            WITH ChildQuery(2)
               .setoDataQuery(_FAMILYSERVTYPES_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Parts in Family", _FAMILY_LOGPART_SUBFORM))
                   .add(1, new oNavigation("Malfunction Codes for Family", _FAMILYMALF_SUBFORM))
                   .add(2, new oNavigation("Service Terms per Family", _FAMILYSERVTYPES_SUBFORM))
                   .add(3, new oNavigation("Required Analyses", _FAMILYREQLABANALYSES_SUBFORM))
               end with
            end with
            WITH ChildQuery(3)
               .setoDataQuery(_FAMILYREQLABANALYSES_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Parts in Family", _FAMILY_LOGPART_SUBFORM))
                   .add(1, new oNavigation("Malfunction Codes for Family", _FAMILYMALF_SUBFORM))
                   .add(2, new oNavigation("Service Terms per Family", _FAMILYSERVTYPES_SUBFORM))
                   .add(3, new oNavigation("Required Analyses", _FAMILYREQLABANALYSES_SUBFORM))
               end with
            end with
        End Sub
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                if _parent is nothing then
                    return "FAMILY_LOG"
                else
                    return "FAMILY_LOG_SUBFORM"
                end if
            End Get
        End Property
        
        Public Overrides ReadOnly Property KeyString() As String
            Get
                return string.format( _
                 "FAMILYNAME={0}", _
                  string.format("'{0}'",FAMILYNAME) _ 
               )
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        <DisplayName("Family"),  _
         nType("Edm.String"),  _
         tab("Family"),  _
         Pos(20),  _
         Mandatory(true),  _
         twodBarcode("FAMILYNAME")>  _
        Public Property FAMILYNAME() As String
            Get
                return _FAMILYNAME
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Family", value, "^.{0,8}$") then Exit Property
                _IsSetFAMILYNAME = True
                If loading Then
                  _FAMILYNAME = Value
                Else
                    if not _FAMILYNAME = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("FAMILYNAME", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _FAMILYNAME = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Family Description"),  _
         nType("Edm.String"),  _
         tab("Family"),  _
         Pos(25),  _
         twodBarcode("FAMILYDESC"),  _
         help("Some help.")>  _
        Public Property FAMILYDESC() As String
            Get
                return _FAMILYDESC
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Family Description", value, "^.{0,32}$") then Exit Property
                _IsSetFAMILYDESC = True
                If loading Then
                  _FAMILYDESC = Value
                Else
                    if not _FAMILYDESC = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("FAMILYDESC", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _FAMILYDESC = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Family Desc (Lang 2)"),  _
         nType("Edm.String"),  _
         tab("Family"),  _
         Pos(30),  _
         twodBarcode("EFAMILYDES")>  _
        Public Property EFAMILYDES() As String
            Get
                return _EFAMILYDES
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Family Desc (Lang 2)", value, "^.{0,32}$") then Exit Property
                _IsSetEFAMILYDES = True
                If loading Then
                  _EFAMILYDES = Value
                Else
                    if not _EFAMILYDES = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("EFAMILYDES", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _EFAMILYDES = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Work Hours?"),  _
         nType("Edm.String"),  _
         tab("Family"),  _
         Pos(31),  _
         twodBarcode("TECHFLAG")>  _
        Public Property TECHFLAG() As String
            Get
                return _TECHFLAG
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Work Hours?", value, "^.{0,1}$") then Exit Property
                _IsSetTECHFLAG = True
                If loading Then
                  _TECHFLAG = Value
                Else
                    if not _TECHFLAG = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("TECHFLAG", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _TECHFLAG = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Bill Project Report?"),  _
         nType("Edm.String"),  _
         tab("Family"),  _
         Pos(32),  _
         twodBarcode("DEBITFLAG")>  _
        Public Property DEBITFLAG() As String
            Get
                return _DEBITFLAG
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Bill Project Report?", value, "^.{0,1}$") then Exit Property
                _IsSetDEBITFLAG = True
                If loading Then
                  _DEBITFLAG = Value
                Else
                    if not _DEBITFLAG = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("DEBITFLAG", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _DEBITFLAG = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Production for Proj?"),  _
         nType("Edm.String"),  _
         tab("Family"),  _
         Pos(35),  _
         twodBarcode("PROJPROCESS")>  _
        Public Property PROJPROCESS() As String
            Get
                return _PROJPROCESS
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Production for Proj?", value, "^.{0,1}$") then Exit Property
                _IsSetPROJPROCESS = True
                If loading Then
                  _PROJPROCESS = Value
                Else
                    if not _PROJPROCESS = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("PROJPROCESS", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _PROJPROCESS = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Technician in Charge"),  _
         nType("Edm.String"),  _
         tab("Family"),  _
         Pos(40),  _
         twodBarcode("TECHNICIANLOGIN")>  _
        Public Property TECHNICIANLOGIN() As String
            Get
                return _TECHNICIANLOGIN
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Technician in Charge", value, "^.{0,20}$") then Exit Property
                _IsSetTECHNICIANLOGIN = True
                If loading Then
                  _TECHNICIANLOGIN = Value
                Else
                    if not _TECHNICIANLOGIN = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("TECHNICIANLOGIN", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _TECHNICIANLOGIN = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Family Type Code"),  _
         nType("Edm.String"),  _
         tab("Family"),  _
         Pos(50),  _
         twodBarcode("FTCODE")>  _
        Public Property FTCODE() As String
            Get
                return _FTCODE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Family Type Code", value, "^.{0,4}$") then Exit Property
                _IsSetFTCODE = True
                If loading Then
                  _FTCODE = Value
                Else
                    if not _FTCODE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("FTCODE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _FTCODE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Family Type Descrip."),  _
         nType("Edm.String"),  _
         tab("Family Type Descrip."),  _
         Pos(52),  _
         [ReadOnly](true),  _
         twodBarcode("FTNAME")>  _
        Public Property FTNAME() As String
            Get
                return _FTNAME
            End Get
            Set
                if not(value is nothing) then
                  _FTNAME = Value
                end if
            End Set
        End Property
        
        <DisplayName("Delivery Scheduling?"),  _
         nType("Edm.String"),  _
         tab("Family Type Descrip."),  _
         Pos(55),  _
         twodBarcode("DISTRFLAG")>  _
        Public Property DISTRFLAG() As String
            Get
                return _DISTRFLAG
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Delivery Scheduling?", value, "^.{0,1}$") then Exit Property
                _IsSetDISTRFLAG = True
                If loading Then
                  _DISTRFLAG = Value
                Else
                    if not _DISTRFLAG = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("DISTRFLAG", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _DISTRFLAG = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Air Cost (%)"),  _
         nType("Edm.Decimal"),  _
         Scale(2),  _
         Precision(6),  _
         tab("Family Type Descrip."),  _
         Pos(60),  _
         twodBarcode("DUTYPERCENT")>  _
        Public Property DUTYPERCENT() As nullable(of decimal)
            Get
                return _DUTYPERCENT
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Air Cost (%)", value, "^[0-9\.\-]+$") then Exit Property
                _IsSetDUTYPERCENT = True
                If loading Then
                  _DUTYPERCENT = Value
                Else
                    if not _DUTYPERCENT = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("DUTYPERCENT", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _DUTYPERCENT = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Sea Cost (%)"),  _
         nType("Edm.Decimal"),  _
         Scale(2),  _
         Precision(6),  _
         tab("Family Type Descrip."),  _
         Pos(70),  _
         twodBarcode("DUTYPERCENT1")>  _
        Public Property DUTYPERCENT1() As nullable(of decimal)
            Get
                return _DUTYPERCENT1
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Sea Cost (%)", value, "^[0-9\.\-]+$") then Exit Property
                _IsSetDUTYPERCENT1 = True
                If loading Then
                  _DUTYPERCENT1 = Value
                Else
                    if not _DUTYPERCENT1 = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("DUTYPERCENT1", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _DUTYPERCENT1 = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Land Cost (%)"),  _
         nType("Edm.Decimal"),  _
         Scale(2),  _
         Precision(6),  _
         tab("Family Type Descrip."),  _
         Pos(80),  _
         twodBarcode("DUTYPERCENT2")>  _
        Public Property DUTYPERCENT2() As nullable(of decimal)
            Get
                return _DUTYPERCENT2
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Land Cost (%)", value, "^[0-9\.\-]+$") then Exit Property
                _IsSetDUTYPERCENT2 = True
                If loading Then
                  _DUTYPERCENT2 = Value
                Else
                    if not _DUTYPERCENT2 = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("DUTYPERCENT2", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _DUTYPERCENT2 = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Shipping Cost Type"),  _
         nType("Edm.String"),  _
         tab("Family Type Descrip."),  _
         Pos(82),  _
         twodBarcode("DUTYPERCENTTYPE")>  _
        Public Property DUTYPERCENTTYPE() As String
            Get
                return _DUTYPERCENTTYPE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Shipping Cost Type", value, "^.{0,1}$") then Exit Property
                _IsSetDUTYPERCENTTYPE = True
                If loading Then
                  _DUTYPERCENTTYPE = Value
                Else
                    if not _DUTYPERCENTTYPE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("DUTYPERCENTTYPE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _DUTYPERCENTTYPE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Sold Part?"),  _
         nType("Edm.String"),  _
         tab("Family Type Descrip."),  _
         Pos(85),  _
         twodBarcode("SELLFLAG")>  _
        Public Property SELLFLAG() As String
            Get
                return _SELLFLAG
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Sold Part?", value, "^.{0,1}$") then Exit Property
                _IsSetSELLFLAG = True
                If loading Then
                  _SELLFLAG = Value
                Else
                    if not _SELLFLAG = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("SELLFLAG", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _SELLFLAG = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Block Neg. Balance?"),  _
         nType("Edm.String"),  _
         tab("Family Type Descrip."),  _
         Pos(90),  _
         twodBarcode("NEGBALFLAG")>  _
        Public Property NEGBALFLAG() As String
            Get
                return _NEGBALFLAG
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Block Neg. Balance?", value, "^.{0,1}$") then Exit Property
                _IsSetNEGBALFLAG = True
                If loading Then
                  _NEGBALFLAG = Value
                Else
                    if not _NEGBALFLAG = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("NEGBALFLAG", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _NEGBALFLAG = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Sales/Order Target?"),  _
         nType("Edm.String"),  _
         tab("Sales/Order Target?"),  _
         Pos(100),  _
         twodBarcode("FORECAST")>  _
        Public Property FORECAST() As String
            Get
                return _FORECAST
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Sales/Order Target?", value, "^.{0,1}$") then Exit Property
                _IsSetFORECAST = True
                If loading Then
                  _FORECAST = Value
                Else
                    if not _FORECAST = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("FORECAST", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _FORECAST = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Sample Part"),  _
         nType("Edm.String"),  _
         tab("Sales/Order Target?"),  _
         Pos(110),  _
         twodBarcode("PARTNAME")>  _
        Public Property PARTNAME() As String
            Get
                return _PARTNAME
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Sample Part", value, "^.{0,15}$") then Exit Property
                _IsSetPARTNAME = True
                If loading Then
                  _PARTNAME = Value
                Else
                    if not _PARTNAME = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("PARTNAME", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _PARTNAME = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Part Description"),  _
         nType("Edm.String"),  _
         tab("Sales/Order Target?"),  _
         Pos(120),  _
         [ReadOnly](true),  _
         twodBarcode("PARTDES")>  _
        Public Property PARTDES() As String
            Get
                return _PARTDES
            End Get
            Set
                if not(value is nothing) then
                  _PARTDES = Value
                end if
            End Set
        End Property
        
        <DisplayName("Put Strategy Code"),  _
         nType("Edm.String"),  _
         tab("Sales/Order Target?"),  _
         Pos(130),  _
         twodBarcode("STORAGETYPECODE")>  _
        Public Property STORAGETYPECODE() As String
            Get
                return _STORAGETYPECODE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Put Strategy Code", value, "^.{0,8}$") then Exit Property
                _IsSetSTORAGETYPECODE = True
                If loading Then
                  _STORAGETYPECODE = Value
                Else
                    if not _STORAGETYPECODE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("STORAGETYPECODE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _STORAGETYPECODE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Put Strategy Descrip"),  _
         nType("Edm.String"),  _
         tab("Sales/Order Target?"),  _
         Pos(140),  _
         [ReadOnly](true),  _
         twodBarcode("STORAGETYPEDES")>  _
        Public Property STORAGETYPEDES() As String
            Get
                return _STORAGETYPEDES
            End Get
            Set
                if not(value is nothing) then
                  _STORAGETYPEDES = Value
                end if
            End Set
        End Property
        
        <DisplayName("Monthly Expir-Pick"),  _
         nType("Edm.String"),  _
         tab("Sales/Order Target?"),  _
         Pos(145),  _
         twodBarcode("SERIALMONTHEXPIRY")>  _
        Public Property SERIALMONTHEXPIRY() As String
            Get
                return _SERIALMONTHEXPIRY
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Monthly Expir-Pick", value, "^.{0,1}$") then Exit Property
                _IsSetSERIALMONTHEXPIRY = True
                If loading Then
                  _SERIALMONTHEXPIRY = Value
                Else
                    if not _SERIALMONTHEXPIRY = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("SERIALMONTHEXPIRY", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _SERIALMONTHEXPIRY = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Top-Level?"),  _
         nType("Edm.String"),  _
         tab("Sales/Order Target?"),  _
         Pos(150),  _
         twodBarcode("ISMFAMILY")>  _
        Public Property ISMFAMILY() As String
            Get
                return _ISMFAMILY
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Top-Level?", value, "^.{0,1}$") then Exit Property
                _IsSetISMFAMILY = True
                If loading Then
                  _ISMFAMILY = Value
                Else
                    if not _ISMFAMILY = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("ISMFAMILY", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _ISMFAMILY = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Top-Level Family"),  _
         nType("Edm.String"),  _
         tab("Sales/Order Target?"),  _
         Pos(165),  _
         twodBarcode("MFAMILYNAME")>  _
        Public Property MFAMILYNAME() As String
            Get
                return _MFAMILYNAME
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Top-Level Family", value, "^.{0,8}$") then Exit Property
                _IsSetMFAMILYNAME = True
                If loading Then
                  _MFAMILYNAME = Value
                Else
                    if not _MFAMILYNAME = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("MFAMILYNAME", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _MFAMILYNAME = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Family Description"),  _
         nType("Edm.String"),  _
         tab("Family Description"),  _
         Pos(170),  _
         [ReadOnly](true),  _
         twodBarcode("MFAMILYDES")>  _
        Public Property MFAMILYDES() As String
            Get
                return _MFAMILYDES
            End Get
            Set
                if not(value is nothing) then
                  _MFAMILYDES = Value
                end if
            End Set
        End Property
        
        <DisplayName("Not in Use"),  _
         nType("Edm.String"),  _
         tab("Family Description"),  _
         Pos(190),  _
         twodBarcode("INACTIVE")>  _
        Public Property INACTIVE() As String
            Get
                return _INACTIVE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Not in Use", value, "^.{0,1}$") then Exit Property
                _IsSetINACTIVE = True
                If loading Then
                  _INACTIVE = Value
                Else
                    if not _INACTIVE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("INACTIVE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _INACTIVE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Family (ID)"),  _
         nType("Edm.Int64"),  _
         tab("Family Description"),  _
         Pos(40),  _
         Browsable(false),  _
         twodBarcode("FAMILY")>  _
        Public Property FAMILY() As nullable (of int64)
            Get
                return _FAMILY
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Family (ID)", value, "^[0-9\-]+$") then Exit Property
                _IsSetFAMILY = True
                If loading Then
                  _FAMILY = Value
                Else
                    if not _FAMILY = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("FAMILY", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _FAMILY = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <JsonIgnore(),  _
         Browsable(false)>  _
        Public Property FAMILY_LOGPART_SUBFORM() As QUERY_FAMILY_LOGPART
            Get
                return _FAMILY_LOGPART_SUBFORM
            End Get
            Set
                _FAMILY_LOGPART_SUBFORM = value
            End Set
        End Property
        
        <JsonIgnore(),  _
         Browsable(false)>  _
        Public Property FAMILYMALF_SUBFORM() As QUERY_FAMILYMALF
            Get
                return _FAMILYMALF_SUBFORM
            End Get
            Set
                _FAMILYMALF_SUBFORM = value
            End Set
        End Property
        
        <JsonIgnore(),  _
         Browsable(false)>  _
        Public Property FAMILYSERVTYPES_SUBFORM() As QUERY_FAMILYSERVTYPES
            Get
                return _FAMILYSERVTYPES_SUBFORM
            End Get
            Set
                _FAMILYSERVTYPES_SUBFORM = value
            End Set
        End Property
        
        <JsonIgnore(),  _
         Browsable(false)>  _
        Public Property FAMILYREQLABANALYSES_SUBFORM() As QUERY_FAMILYREQLABANALYSES
            Get
                return _FAMILYREQLABANALYSES_SUBFORM
            End Get
            Set
                _FAMILYREQLABANALYSES_SUBFORM = value
            End Set
        End Property
        
        Protected Friend Overrides Sub toJson(ByRef jw As Newtonsoft.Json.JsonTextWriter)
            Dim f as boolean = false
            if _IsSetFAMILYNAME then
              if f then
                  jw.WriteRaw(", ""FAMILYNAME"": ")
              else
                  jw.WriteRaw("""FAMILYNAME"": ")
                  f = true
              end if
              jw.WriteValue(me.FAMILYNAME)
            end if
            if _IsSetFAMILYDESC then
              if f then
                  jw.WriteRaw(", ""FAMILYDESC"": ")
              else
                  jw.WriteRaw("""FAMILYDESC"": ")
                  f = true
              end if
              jw.WriteValue(me.FAMILYDESC)
            end if
            if _IsSetEFAMILYDES then
              if f then
                  jw.WriteRaw(", ""EFAMILYDES"": ")
              else
                  jw.WriteRaw("""EFAMILYDES"": ")
                  f = true
              end if
              jw.WriteValue(me.EFAMILYDES)
            end if
            if _IsSetTECHFLAG then
              if f then
                  jw.WriteRaw(", ""TECHFLAG"": ")
              else
                  jw.WriteRaw("""TECHFLAG"": ")
                  f = true
              end if
              jw.WriteValue(me.TECHFLAG)
            end if
            if _IsSetDEBITFLAG then
              if f then
                  jw.WriteRaw(", ""DEBITFLAG"": ")
              else
                  jw.WriteRaw("""DEBITFLAG"": ")
                  f = true
              end if
              jw.WriteValue(me.DEBITFLAG)
            end if
            if _IsSetPROJPROCESS then
              if f then
                  jw.WriteRaw(", ""PROJPROCESS"": ")
              else
                  jw.WriteRaw("""PROJPROCESS"": ")
                  f = true
              end if
              jw.WriteValue(me.PROJPROCESS)
            end if
            if _IsSetTECHNICIANLOGIN then
              if f then
                  jw.WriteRaw(", ""TECHNICIANLOGIN"": ")
              else
                  jw.WriteRaw("""TECHNICIANLOGIN"": ")
                  f = true
              end if
              jw.WriteValue(me.TECHNICIANLOGIN)
            end if
            if _IsSetFTCODE then
              if f then
                  jw.WriteRaw(", ""FTCODE"": ")
              else
                  jw.WriteRaw("""FTCODE"": ")
                  f = true
              end if
              jw.WriteValue(me.FTCODE)
            end if
            if _IsSetDISTRFLAG then
              if f then
                  jw.WriteRaw(", ""DISTRFLAG"": ")
              else
                  jw.WriteRaw("""DISTRFLAG"": ")
                  f = true
              end if
              jw.WriteValue(me.DISTRFLAG)
            end if
            if _IsSetDUTYPERCENT then
              if f then
                  jw.WriteRaw(", ""DUTYPERCENT"": ")
              else
                  jw.WriteRaw("""DUTYPERCENT"": ")
                  f = true
              end if
              jw.WriteValue(me.DUTYPERCENT)
            end if
            if _IsSetDUTYPERCENT1 then
              if f then
                  jw.WriteRaw(", ""DUTYPERCENT1"": ")
              else
                  jw.WriteRaw("""DUTYPERCENT1"": ")
                  f = true
              end if
              jw.WriteValue(me.DUTYPERCENT1)
            end if
            if _IsSetDUTYPERCENT2 then
              if f then
                  jw.WriteRaw(", ""DUTYPERCENT2"": ")
              else
                  jw.WriteRaw("""DUTYPERCENT2"": ")
                  f = true
              end if
              jw.WriteValue(me.DUTYPERCENT2)
            end if
            if _IsSetDUTYPERCENTTYPE then
              if f then
                  jw.WriteRaw(", ""DUTYPERCENTTYPE"": ")
              else
                  jw.WriteRaw("""DUTYPERCENTTYPE"": ")
                  f = true
              end if
              jw.WriteValue(me.DUTYPERCENTTYPE)
            end if
            if _IsSetSELLFLAG then
              if f then
                  jw.WriteRaw(", ""SELLFLAG"": ")
              else
                  jw.WriteRaw("""SELLFLAG"": ")
                  f = true
              end if
              jw.WriteValue(me.SELLFLAG)
            end if
            if _IsSetNEGBALFLAG then
              if f then
                  jw.WriteRaw(", ""NEGBALFLAG"": ")
              else
                  jw.WriteRaw("""NEGBALFLAG"": ")
                  f = true
              end if
              jw.WriteValue(me.NEGBALFLAG)
            end if
            if _IsSetFORECAST then
              if f then
                  jw.WriteRaw(", ""FORECAST"": ")
              else
                  jw.WriteRaw("""FORECAST"": ")
                  f = true
              end if
              jw.WriteValue(me.FORECAST)
            end if
            if _IsSetPARTNAME then
              if f then
                  jw.WriteRaw(", ""PARTNAME"": ")
              else
                  jw.WriteRaw("""PARTNAME"": ")
                  f = true
              end if
              jw.WriteValue(me.PARTNAME)
            end if
            if _IsSetSTORAGETYPECODE then
              if f then
                  jw.WriteRaw(", ""STORAGETYPECODE"": ")
              else
                  jw.WriteRaw("""STORAGETYPECODE"": ")
                  f = true
              end if
              jw.WriteValue(me.STORAGETYPECODE)
            end if
            if _IsSetSERIALMONTHEXPIRY then
              if f then
                  jw.WriteRaw(", ""SERIALMONTHEXPIRY"": ")
              else
                  jw.WriteRaw("""SERIALMONTHEXPIRY"": ")
                  f = true
              end if
              jw.WriteValue(me.SERIALMONTHEXPIRY)
            end if
            if _IsSetISMFAMILY then
              if f then
                  jw.WriteRaw(", ""ISMFAMILY"": ")
              else
                  jw.WriteRaw("""ISMFAMILY"": ")
                  f = true
              end if
              jw.WriteValue(me.ISMFAMILY)
            end if
            if _IsSetMFAMILYNAME then
              if f then
                  jw.WriteRaw(", ""MFAMILYNAME"": ")
              else
                  jw.WriteRaw("""MFAMILYNAME"": ")
                  f = true
              end if
              jw.WriteValue(me.MFAMILYNAME)
            end if
            if _IsSetINACTIVE then
              if f then
                  jw.WriteRaw(", ""INACTIVE"": ")
              else
                  jw.WriteRaw("""INACTIVE"": ")
                  f = true
              end if
              jw.WriteValue(me.INACTIVE)
            end if
            if _IsSetFAMILY then
              if f then
                  jw.WriteRaw(", ""FAMILY"": ")
              else
                  jw.WriteRaw("""FAMILY"": ")
                  f = true
              end if
              jw.WriteValue(me.FAMILY)
            end if
            if _FAMILY_LOGPART_SUBFORM.value.count > 0 then
              jw.WriteRaw(", FAMILY_LOGPART_SUBFORM: ")
              jw.WriteRaw("[ ")
              dim i as integer = 0
              for each itm as FAMILY_LOGPART in _FAMILY_LOGPART_SUBFORM.value
                jw.WriteRaw("{ ")
                itm.tojson(jw)
                jw.WriteRaw(" }")
                if i < _FAMILY_LOGPART_SUBFORM.value.count - 1 then jw.WriteRaw(", ")
                i += 1
                next
              jw.WriteRaw(" ]")
            end if
            if _FAMILYMALF_SUBFORM.value.count > 0 then
              jw.WriteRaw(", FAMILYMALF_SUBFORM: ")
              jw.WriteRaw("[ ")
              dim i as integer = 0
              for each itm as FAMILYMALF in _FAMILYMALF_SUBFORM.value
                jw.WriteRaw("{ ")
                itm.tojson(jw)
                jw.WriteRaw(" }")
                if i < _FAMILYMALF_SUBFORM.value.count - 1 then jw.WriteRaw(", ")
                i += 1
                next
              jw.WriteRaw(" ]")
            end if
            if _FAMILYSERVTYPES_SUBFORM.value.count > 0 then
              jw.WriteRaw(", FAMILYSERVTYPES_SUBFORM: ")
              jw.WriteRaw("[ ")
              dim i as integer = 0
              for each itm as FAMILYSERVTYPES in _FAMILYSERVTYPES_SUBFORM.value
                jw.WriteRaw("{ ")
                itm.tojson(jw)
                jw.WriteRaw(" }")
                if i < _FAMILYSERVTYPES_SUBFORM.value.count - 1 then jw.WriteRaw(", ")
                i += 1
                next
              jw.WriteRaw(" ]")
            end if
            if _FAMILYREQLABANALYSES_SUBFORM.value.count > 0 then
              jw.WriteRaw(", FAMILYREQLABANALYSES_SUBFORM: ")
              jw.WriteRaw("[ ")
              dim i as integer = 0
              for each itm as FAMILYREQLABANALYSES in _FAMILYREQLABANALYSES_SUBFORM.value
                jw.WriteRaw("{ ")
                itm.tojson(jw)
                jw.WriteRaw(" }")
                if i < _FAMILYREQLABANALYSES_SUBFORM.value.count - 1 then jw.WriteRaw(", ")
                i += 1
                next
              jw.WriteRaw(" ]")
            end if
        End Sub
        
        Protected Friend Overrides Sub toXML(ByRef xw As System.Xml.XmlWriter, ByVal name As String)
            with xw
              .WriteStartElement("form")
              if name is nothing then
                .WriteAttributeString("name", "FAMILY_LOG")
              else
                .WriteAttributeString("name", name)
              end if
              .WriteAttributeString("result", 0)
              .WriteStartElement("key")
              .WriteAttributeString("name", "FAMILYNAME")
              .WriteAttributeString("value", "")
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "8")
              .WriteEndElement
            if _IsSetFAMILYNAME then
              .WriteStartElement("field")
              .WriteAttributeString("name", "FAMILYNAME")
              .WriteAttributeString("value", me.FAMILYNAME)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "8")
              .WriteEndElement
            end if
            if _IsSetFAMILYDESC then
              .WriteStartElement("field")
              .WriteAttributeString("name", "FAMILYDESC")
              .WriteAttributeString("value", me.FAMILYDESC)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "32")
              .WriteEndElement
            end if
            if _IsSetEFAMILYDES then
              .WriteStartElement("field")
              .WriteAttributeString("name", "EFAMILYDES")
              .WriteAttributeString("value", me.EFAMILYDES)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "32")
              .WriteEndElement
            end if
            if _IsSetTECHFLAG then
              .WriteStartElement("field")
              .WriteAttributeString("name", "TECHFLAG")
              .WriteAttributeString("value", me.TECHFLAG)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetDEBITFLAG then
              .WriteStartElement("field")
              .WriteAttributeString("name", "DEBITFLAG")
              .WriteAttributeString("value", me.DEBITFLAG)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetPROJPROCESS then
              .WriteStartElement("field")
              .WriteAttributeString("name", "PROJPROCESS")
              .WriteAttributeString("value", me.PROJPROCESS)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetTECHNICIANLOGIN then
              .WriteStartElement("field")
              .WriteAttributeString("name", "TECHNICIANLOGIN")
              .WriteAttributeString("value", me.TECHNICIANLOGIN)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "20")
              .WriteEndElement
            end if
            if _IsSetFTCODE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "FTCODE")
              .WriteAttributeString("value", me.FTCODE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "4")
              .WriteEndElement
            end if
            if _IsSetDISTRFLAG then
              .WriteStartElement("field")
              .WriteAttributeString("name", "DISTRFLAG")
              .WriteAttributeString("value", me.DISTRFLAG)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetDUTYPERCENT then
              .WriteStartElement("field")
              .WriteAttributeString("name", "DUTYPERCENT")
              .WriteAttributeString("value", me.DUTYPERCENT)
              .WriteAttributeString("type", "Edm.Decimal")
              .WriteAttributeString("Scale", "2")
              .WriteAttributeString("Precision", "6")
              .WriteEndElement
            end if
            if _IsSetDUTYPERCENT1 then
              .WriteStartElement("field")
              .WriteAttributeString("name", "DUTYPERCENT1")
              .WriteAttributeString("value", me.DUTYPERCENT1)
              .WriteAttributeString("type", "Edm.Decimal")
              .WriteAttributeString("Scale", "2")
              .WriteAttributeString("Precision", "6")
              .WriteEndElement
            end if
            if _IsSetDUTYPERCENT2 then
              .WriteStartElement("field")
              .WriteAttributeString("name", "DUTYPERCENT2")
              .WriteAttributeString("value", me.DUTYPERCENT2)
              .WriteAttributeString("type", "Edm.Decimal")
              .WriteAttributeString("Scale", "2")
              .WriteAttributeString("Precision", "6")
              .WriteEndElement
            end if
            if _IsSetDUTYPERCENTTYPE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "DUTYPERCENTTYPE")
              .WriteAttributeString("value", me.DUTYPERCENTTYPE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetSELLFLAG then
              .WriteStartElement("field")
              .WriteAttributeString("name", "SELLFLAG")
              .WriteAttributeString("value", me.SELLFLAG)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetNEGBALFLAG then
              .WriteStartElement("field")
              .WriteAttributeString("name", "NEGBALFLAG")
              .WriteAttributeString("value", me.NEGBALFLAG)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetFORECAST then
              .WriteStartElement("field")
              .WriteAttributeString("name", "FORECAST")
              .WriteAttributeString("value", me.FORECAST)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetPARTNAME then
              .WriteStartElement("field")
              .WriteAttributeString("name", "PARTNAME")
              .WriteAttributeString("value", me.PARTNAME)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "15")
              .WriteEndElement
            end if
            if _IsSetSTORAGETYPECODE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "STORAGETYPECODE")
              .WriteAttributeString("value", me.STORAGETYPECODE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "8")
              .WriteEndElement
            end if
            if _IsSetSERIALMONTHEXPIRY then
              .WriteStartElement("field")
              .WriteAttributeString("name", "SERIALMONTHEXPIRY")
              .WriteAttributeString("value", me.SERIALMONTHEXPIRY)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetISMFAMILY then
              .WriteStartElement("field")
              .WriteAttributeString("name", "ISMFAMILY")
              .WriteAttributeString("value", me.ISMFAMILY)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetMFAMILYNAME then
              .WriteStartElement("field")
              .WriteAttributeString("name", "MFAMILYNAME")
              .WriteAttributeString("value", me.MFAMILYNAME)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "8")
              .WriteEndElement
            end if
            if _IsSetINACTIVE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "INACTIVE")
              .WriteAttributeString("value", me.INACTIVE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetFAMILY then
              .WriteStartElement("field")
              .WriteAttributeString("name", "FAMILY")
              .WriteAttributeString("value", me.FAMILY)
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
            end if
            if _FAMILY_LOGPART_SUBFORM.value.count > 0 then
              for each itm as FAMILY_LOGPART in _FAMILY_LOGPART_SUBFORM.Value
                itm.toXML(xw,"FAMILY_LOGPART_SUBFORM")
              next
            end if
            if _FAMILYMALF_SUBFORM.value.count > 0 then
              for each itm as FAMILYMALF in _FAMILYMALF_SUBFORM.Value
                itm.toXML(xw,"FAMILYMALF_SUBFORM")
              next
            end if
            if _FAMILYSERVTYPES_SUBFORM.value.count > 0 then
              for each itm as FAMILYSERVTYPES in _FAMILYSERVTYPES_SUBFORM.Value
                itm.toXML(xw,"FAMILYSERVTYPES_SUBFORM")
              next
            end if
            if _FAMILYREQLABANALYSES_SUBFORM.value.count > 0 then
              for each itm as FAMILYREQLABANALYSES in _FAMILYREQLABANALYSES_SUBFORM.Value
                itm.toXML(xw,"FAMILYREQLABANALYSES_SUBFORM")
              next
            end if
              .WriteEndElement
            end with
        End Sub
        
        Protected Friend Overrides Sub HandlesEdit(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILY_LOG = JsonConvert.DeserializeObject(Of FAMILY_LOG)(e.StreamReader.ReadToEnd)
                With obj
                  _FAMILYNAME = .FAMILYNAME
                  _FAMILYDESC = .FAMILYDESC
                  _EFAMILYDES = .EFAMILYDES
                  _TECHFLAG = .TECHFLAG
                  _DEBITFLAG = .DEBITFLAG
                  _PROJPROCESS = .PROJPROCESS
                  _TECHNICIANLOGIN = .TECHNICIANLOGIN
                  _FTCODE = .FTCODE
                  _FTNAME = .FTNAME
                  _DISTRFLAG = .DISTRFLAG
                  _DUTYPERCENT = .DUTYPERCENT
                  _DUTYPERCENT1 = .DUTYPERCENT1
                  _DUTYPERCENT2 = .DUTYPERCENT2
                  _DUTYPERCENTTYPE = .DUTYPERCENTTYPE
                  _SELLFLAG = .SELLFLAG
                  _NEGBALFLAG = .NEGBALFLAG
                  _FORECAST = .FORECAST
                  _PARTNAME = .PARTNAME
                  _PARTDES = .PARTDES
                  _STORAGETYPECODE = .STORAGETYPECODE
                  _STORAGETYPEDES = .STORAGETYPEDES
                  _SERIALMONTHEXPIRY = .SERIALMONTHEXPIRY
                  _ISMFAMILY = .ISMFAMILY
                  _MFAMILYNAME = .MFAMILYNAME
                  _MFAMILYDES = .MFAMILYDES
                  _INACTIVE = .INACTIVE
                  _FAMILY = .FAMILY
                end with
            End If
        End Sub
    End Class
    
    Public Enum Subform_FAMILY_LOG
        
        FAMILY_LOGPART = 0
        
        FAMILYMALF = 1
        
        FAMILYSERVTYPES = 2
        
        FAMILYREQLABANALYSES = 3
    End Enum
    
    <QueryTitle("Parts in Family")>  _
    Public Class QUERY_FAMILY_LOGPART
        Inherits oDataQuery
        
        Private _Name As String
        
        Private _Value As list(of FAMILY_LOGPART)
        
        Private _Parent As oDataObject
        
        Private _BindingSource As System.Windows.Forms.BindingSource
        
        Public Sub New()
            MyBase.New
            _Value = new list (of FAMILY_LOGPART)
            _Parent = nothing
            _Name = "FAMILY_LOGPART"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            end with
        End Sub
        
        Public Sub New(ByRef Parent As oDataObject)
            MyBase.New
            _Value = new list (of FAMILY_LOGPART)
            _Parent = Parent
            _name = "FAMILY_LOGPART_SUBFORM"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            end with
        End Sub
        
        <JsonProperty("@odata.context")>  _
        Protected Friend Property Name() As String
            Get
                return _Name
            End Get
            Set
                _Name = value
            End Set
        End Property
        
        Public Property Value() As list(of FAMILY_LOGPART)
            Get
                return _Value
            End Get
            Set
                _Value = value
            End Set
        End Property
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                return _Name
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        Public Overrides ReadOnly Property ObjectType() As System.Type
            Get
                return GetType(FAMILY_LOGPART)
            End Get
        End Property
        
        Public Overrides ReadOnly Property BindingSource() As System.Windows.Forms.BindingSource
            Get
                Return _BindingSource
            End Get
        End Property
        
        Protected Friend Overrides Sub Deserialise(ByRef Stream As StreamReader)
            _value.clear
            For Each _FAMILY_LOGPART As FAMILY_LOGPART In JsonConvert.DeserializeObject(Of QUERY_FAMILY_LOGPART)(stream.ReadToEnd).Value
              With _FAMILY_LOGPART
                  .Parent = Parent
                  .loading = false
              end with
              _value.add(_FAMILY_LOGPART)
            next
            _BindingSource.DataSource = _Value
        End Sub
        
        Protected Friend Overrides Sub HandlesAdd(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILY_LOGPART = JsonConvert.DeserializeObject(Of FAMILY_LOGPART)(e.StreamReader.ReadToEnd)
                With TryCast(BindingSource.Current, FAMILY_LOGPART)
                  .PARTNAME = obj.PARTNAME
                  .TYPE = obj.TYPE
                  .PARTDES = obj.PARTDES
                  .STATDES = obj.STATDES
                  .PART = obj.PART
                  .loading = false
                end with
            End If
        End Sub
        
        Private Sub HandlesBeginAdd(ByVal sender As Object, ByVal e As ComponentModel.AddingNewEventArgs)
            e.NewObject = new FAMILY_LOGPART(me.Parent)
        End Sub
        
        Protected Friend Overrides Sub Remove(ByRef obj As oDataObject)
            for each _FAMILY_LOGPART as FAMILY_LOGPART in value
              If _FAMILY_LOGPART.Equals(trycast(obj,FAMILY_LOGPART)) Then
                  value.remove(_FAMILY_LOGPART)
                  exit for
              end if
            next
        End Sub
    End Class
    
    Public Class FAMILY_LOGPART
        Inherits oDataObject
        
        Private _Parent As oDataObject
        
        Private _PARTNAME As String
        
        Private _TYPE As String
        
        Private _PARTDES As String
        
        Private _STATDES As String
        
        Private _PART As Long
        
        Public Sub New(ByVal Parent As oDataObject)
            MyBase.New
            _parent = Parent
        End Sub
        
        Public Sub New()
            MyBase.New
        End Sub
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                if _parent is nothing then
                    return "FAMILY_LOGPART"
                else
                    return "FAMILY_LOGPART_SUBFORM"
                end if
            End Get
        End Property
        
        Public Overrides ReadOnly Property KeyString() As String
            Get
                return string.format( _
                 "PARTNAME={0}", _
                  string.format("'{0}'",PARTNAME) _ 
               )
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        <DisplayName("Part Number"),  _
         nType("Edm.String"),  _
         tab("Part Number"),  _
         Pos(10),  _
         [ReadOnly](true),  _
         twodBarcode("PARTNAME")>  _
        Public Property PARTNAME() As String
            Get
                return _PARTNAME
            End Get
            Set
                if not(value is nothing) then
                  _PARTNAME = Value
                end if
            End Set
        End Property
        
        <DisplayName("Type (P/R/O)"),  _
         nType("Edm.String"),  _
         tab("Part Number"),  _
         Pos(20),  _
         [ReadOnly](true),  _
         twodBarcode("TYPE")>  _
        Public Property TYPE() As String
            Get
                return _TYPE
            End Get
            Set
                if not(value is nothing) then
                  _TYPE = Value
                end if
            End Set
        End Property
        
        <DisplayName("Part Description"),  _
         nType("Edm.String"),  _
         tab("Part Number"),  _
         Pos(50),  _
         [ReadOnly](true),  _
         twodBarcode("PARTDES")>  _
        Public Property PARTDES() As String
            Get
                return _PARTDES
            End Get
            Set
                if not(value is nothing) then
                  _PARTDES = Value
                end if
            End Set
        End Property
        
        <DisplayName("Part Status"),  _
         nType("Edm.String"),  _
         tab("Part Number"),  _
         Pos(60),  _
         [ReadOnly](true),  _
         twodBarcode("STATDES")>  _
        Public Property STATDES() As String
            Get
                return _STATDES
            End Get
            Set
                if not(value is nothing) then
                  _STATDES = Value
                end if
            End Set
        End Property
        
        <DisplayName("Part (ID)"),  _
         nType("Edm.Int64"),  _
         tab("Part Number"),  _
         Pos(20),  _
         [ReadOnly](true),  _
         Browsable(false),  _
         twodBarcode("PART")>  _
        Public Property PART() As nullable (of int64)
            Get
                return _PART
            End Get
            Set
                if not(value is nothing) then
                  _PART = Value
                end if
            End Set
        End Property
        
        Protected Friend Overrides Sub toJson(ByRef jw As Newtonsoft.Json.JsonTextWriter)
            Dim f as boolean = false
        End Sub
        
        Protected Friend Overrides Sub toXML(ByRef xw As System.Xml.XmlWriter, ByVal name As String)
            with xw
              .WriteStartElement("form")
              if name is nothing then
                .WriteAttributeString("name", "FAMILY_LOGPART")
              else
                .WriteAttributeString("name", name)
              end if
              .WriteAttributeString("result", 0)
              .WriteStartElement("key")
              .WriteAttributeString("name", "PARTNAME")
              .WriteAttributeString("value", "")
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "15")
              .WriteEndElement
              .WriteEndElement
            end with
        End Sub
        
        Protected Friend Overrides Sub HandlesEdit(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILY_LOGPART = JsonConvert.DeserializeObject(Of FAMILY_LOGPART)(e.StreamReader.ReadToEnd)
                With obj
                  _PARTNAME = .PARTNAME
                  _TYPE = .TYPE
                  _PARTDES = .PARTDES
                  _STATDES = .STATDES
                  _PART = .PART
                end with
            End If
        End Sub
    End Class
    
    <QueryTitle("Malfunction Codes for Family")>  _
    Public Class QUERY_FAMILYMALF
        Inherits oDataQuery
        
        Private _Name As String
        
        Private _Value As list(of FAMILYMALF)
        
        Private _Parent As oDataObject
        
        Private _BindingSource As System.Windows.Forms.BindingSource
        
        Public Sub New()
            MyBase.New
            _Value = new list (of FAMILYMALF)
            _Parent = nothing
            _Name = "FAMILYMALF"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            end with
        End Sub
        
        Public Sub New(ByRef Parent As oDataObject)
            MyBase.New
            _Value = new list (of FAMILYMALF)
            _Parent = Parent
            _name = "FAMILYMALF_SUBFORM"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            end with
        End Sub
        
        <JsonProperty("@odata.context")>  _
        Protected Friend Property Name() As String
            Get
                return _Name
            End Get
            Set
                _Name = value
            End Set
        End Property
        
        Public Property Value() As list(of FAMILYMALF)
            Get
                return _Value
            End Get
            Set
                _Value = value
            End Set
        End Property
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                return _Name
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        Public Overrides ReadOnly Property ObjectType() As System.Type
            Get
                return GetType(FAMILYMALF)
            End Get
        End Property
        
        Public Overrides ReadOnly Property BindingSource() As System.Windows.Forms.BindingSource
            Get
                Return _BindingSource
            End Get
        End Property
        
        Protected Friend Overrides Sub Deserialise(ByRef Stream As StreamReader)
            _value.clear
            For Each _FAMILYMALF As FAMILYMALF In JsonConvert.DeserializeObject(Of QUERY_FAMILYMALF)(stream.ReadToEnd).Value
              With _FAMILYMALF
                  .Parent = Parent
                  .loading = false
              end with
              _value.add(_FAMILYMALF)
            next
            _BindingSource.DataSource = _Value
        End Sub
        
        Protected Friend Overrides Sub HandlesAdd(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILYMALF = JsonConvert.DeserializeObject(Of FAMILYMALF)(e.StreamReader.ReadToEnd)
                With TryCast(BindingSource.Current, FAMILYMALF)
                  .MALFCODE = obj.MALFCODE
                  .MALFDES = obj.MALFDES
                  .MALF = obj.MALF
                  .loading = false
                end with
            End If
        End Sub
        
        Private Sub HandlesBeginAdd(ByVal sender As Object, ByVal e As ComponentModel.AddingNewEventArgs)
            e.NewObject = new FAMILYMALF(me.Parent)
        End Sub
        
        Protected Friend Overrides Sub Remove(ByRef obj As oDataObject)
            for each _FAMILYMALF as FAMILYMALF in value
              If _FAMILYMALF.Equals(trycast(obj,FAMILYMALF)) Then
                  value.remove(_FAMILYMALF)
                  exit for
              end if
            next
        End Sub
    End Class
    
    Public Class FAMILYMALF
        Inherits oDataObject
        
        Private _Parent As oDataObject
        
        Private _IsSetMALFCODE As Boolean = Boolean.FalseString
        
        Private _MALFCODE As String
        
        Private _MALFDES As String
        
        Private _IsSetMALF As Boolean = Boolean.FalseString
        
        Private _MALF As Long
        
        Public Sub New(ByVal Parent As oDataObject)
            MyBase.New
            _parent = Parent
        End Sub
        
        Public Sub New()
            MyBase.New
        End Sub
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                if _parent is nothing then
                    return "FAMILYMALF"
                else
                    return "FAMILYMALF_SUBFORM"
                end if
            End Get
        End Property
        
        Public Overrides ReadOnly Property KeyString() As String
            Get
                return string.format( _
                 "MALF={0}", _
                  string.format("{0}",MALF) _ 
               )
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        <DisplayName("Malfunction Code"),  _
         nType("Edm.String"),  _
         tab("Malfunction Code"),  _
         Pos(30),  _
         twodBarcode("MALFCODE")>  _
        Public Property MALFCODE() As String
            Get
                return _MALFCODE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Malfunction Code", value, "^.{0,3}$") then Exit Property
                _IsSetMALFCODE = True
                If loading Then
                  _MALFCODE = Value
                Else
                    if not _MALFCODE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("MALFCODE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _MALFCODE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Malf. Description"),  _
         nType("Edm.String"),  _
         tab("Malfunction Code"),  _
         Pos(40),  _
         [ReadOnly](true),  _
         twodBarcode("MALFDES")>  _
        Public Property MALFDES() As String
            Get
                return _MALFDES
            End Get
            Set
                if not(value is nothing) then
                  _MALFDES = Value
                end if
            End Set
        End Property
        
        <DisplayName("Malfunction Code(ID)"),  _
         nType("Edm.Int64"),  _
         tab("Malfunction Code"),  _
         Pos(20),  _
         Browsable(false),  _
         twodBarcode("MALF")>  _
        Public Property MALF() As nullable (of int64)
            Get
                return _MALF
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Malfunction Code(ID)", value, "^[0-9\-]+$") then Exit Property
                _IsSetMALF = True
                If loading Then
                  _MALF = Value
                Else
                    if not _MALF = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("MALF", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _MALF = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        Protected Friend Overrides Sub toJson(ByRef jw As Newtonsoft.Json.JsonTextWriter)
            Dim f as boolean = false
            if _IsSetMALFCODE then
              if f then
                  jw.WriteRaw(", ""MALFCODE"": ")
              else
                  jw.WriteRaw("""MALFCODE"": ")
                  f = true
              end if
              jw.WriteValue(me.MALFCODE)
            end if
            if _IsSetMALF then
              if f then
                  jw.WriteRaw(", ""MALF"": ")
              else
                  jw.WriteRaw("""MALF"": ")
                  f = true
              end if
              jw.WriteValue(me.MALF)
            end if
        End Sub
        
        Protected Friend Overrides Sub toXML(ByRef xw As System.Xml.XmlWriter, ByVal name As String)
            with xw
              .WriteStartElement("form")
              if name is nothing then
                .WriteAttributeString("name", "FAMILYMALF")
              else
                .WriteAttributeString("name", name)
              end if
              .WriteAttributeString("result", 0)
              .WriteStartElement("key")
              .WriteAttributeString("name", "MALF")
              .WriteAttributeString("value", "")
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
            if _IsSetMALFCODE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "MALFCODE")
              .WriteAttributeString("value", me.MALFCODE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "3")
              .WriteEndElement
            end if
            if _IsSetMALF then
              .WriteStartElement("field")
              .WriteAttributeString("name", "MALF")
              .WriteAttributeString("value", me.MALF)
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
            end if
              .WriteEndElement
            end with
        End Sub
        
        Protected Friend Overrides Sub HandlesEdit(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILYMALF = JsonConvert.DeserializeObject(Of FAMILYMALF)(e.StreamReader.ReadToEnd)
                With obj
                  _MALFCODE = .MALFCODE
                  _MALFDES = .MALFDES
                  _MALF = .MALF
                end with
            End If
        End Sub
    End Class
    
    <QueryTitle("Service Terms per Family")>  _
    Public Class QUERY_FAMILYSERVTYPES
        Inherits oDataQuery
        
        Private _Name As String
        
        Private _Value As list(of FAMILYSERVTYPES)
        
        Private _Parent As oDataObject
        
        Private _BindingSource As System.Windows.Forms.BindingSource
        
        Public Sub New()
            MyBase.New
            _Value = new list (of FAMILYSERVTYPES)
            _Parent = nothing
            _Name = "FAMILYSERVTYPES"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            end with
        End Sub
        
        Public Sub New(ByRef Parent As oDataObject)
            MyBase.New
            _Value = new list (of FAMILYSERVTYPES)
            _Parent = Parent
            _name = "FAMILYSERVTYPES_SUBFORM"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            end with
        End Sub
        
        <JsonProperty("@odata.context")>  _
        Protected Friend Property Name() As String
            Get
                return _Name
            End Get
            Set
                _Name = value
            End Set
        End Property
        
        Public Property Value() As list(of FAMILYSERVTYPES)
            Get
                return _Value
            End Get
            Set
                _Value = value
            End Set
        End Property
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                return _Name
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        Public Overrides ReadOnly Property ObjectType() As System.Type
            Get
                return GetType(FAMILYSERVTYPES)
            End Get
        End Property
        
        Public Overrides ReadOnly Property BindingSource() As System.Windows.Forms.BindingSource
            Get
                Return _BindingSource
            End Get
        End Property
        
        Protected Friend Overrides Sub Deserialise(ByRef Stream As StreamReader)
            _value.clear
            For Each _FAMILYSERVTYPES As FAMILYSERVTYPES In JsonConvert.DeserializeObject(Of QUERY_FAMILYSERVTYPES)(stream.ReadToEnd).Value
              With _FAMILYSERVTYPES
                  .Parent = Parent
                  .loading = false
              end with
              _value.add(_FAMILYSERVTYPES)
            next
            _BindingSource.DataSource = _Value
        End Sub
        
        Protected Friend Overrides Sub HandlesAdd(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILYSERVTYPES = JsonConvert.DeserializeObject(Of FAMILYSERVTYPES)(e.StreamReader.ReadToEnd)
                With TryCast(BindingSource.Current, FAMILYSERVTYPES)
                  .SERVTCODE = obj.SERVTCODE
                  .SERVTDES = obj.SERVTDES
                  .FLAG = obj.FLAG
                  .NOCHARGENAME = obj.NOCHARGENAME
                  .SERVTYPE = obj.SERVTYPE
                  .loading = false
                end with
            End If
        End Sub
        
        Private Sub HandlesBeginAdd(ByVal sender As Object, ByVal e As ComponentModel.AddingNewEventArgs)
            e.NewObject = new FAMILYSERVTYPES(me.Parent)
        End Sub
        
        Protected Friend Overrides Sub Remove(ByRef obj As oDataObject)
            for each _FAMILYSERVTYPES as FAMILYSERVTYPES in value
              If _FAMILYSERVTYPES.Equals(trycast(obj,FAMILYSERVTYPES)) Then
                  value.remove(_FAMILYSERVTYPES)
                  exit for
              end if
            next
        End Sub
    End Class
    
    Public Class FAMILYSERVTYPES
        Inherits oDataObject
        
        Private _Parent As oDataObject
        
        Private _IsSetSERVTCODE As Boolean = Boolean.FalseString
        
        Private _SERVTCODE As String
        
        Private _SERVTDES As String
        
        Private _IsSetFLAG As Boolean = Boolean.FalseString
        
        Private _FLAG As String
        
        Private _IsSetNOCHARGENAME As Boolean = Boolean.FalseString
        
        Private _NOCHARGENAME As String
        
        Private _IsSetSERVTYPE As Boolean = Boolean.FalseString
        
        Private _SERVTYPE As Long
        
        Public Sub New(ByVal Parent As oDataObject)
            MyBase.New
            _parent = Parent
        End Sub
        
        Public Sub New()
            MyBase.New
        End Sub
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                if _parent is nothing then
                    return "FAMILYSERVTYPES"
                else
                    return "FAMILYSERVTYPES_SUBFORM"
                end if
            End Get
        End Property
        
        Public Overrides ReadOnly Property KeyString() As String
            Get
                return string.format( _
                 "SERVTYPE={0}", _
                  string.format("{0}",SERVTYPE) _ 
               )
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        <DisplayName("Service Terms (Code)"),  _
         nType("Edm.String"),  _
         tab("Service Terms (Code)"),  _
         Pos(10),  _
         twodBarcode("SERVTCODE")>  _
        Public Property SERVTCODE() As String
            Get
                return _SERVTCODE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Service Terms (Code)", value, "^.{0,3}$") then Exit Property
                _IsSetSERVTCODE = True
                If loading Then
                  _SERVTCODE = Value
                Else
                    if not _SERVTCODE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("SERVTCODE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _SERVTCODE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Service Terms"),  _
         nType("Edm.String"),  _
         tab("Service Terms (Code)"),  _
         Pos(12),  _
         [ReadOnly](true),  _
         twodBarcode("SERVTDES")>  _
        Public Property SERVTDES() As String
            Get
                return _SERVTDES
            End Get
            Set
                if not(value is nothing) then
                  _SERVTDES = Value
                end if
            End Set
        End Property
        
        <DisplayName("Billable?"),  _
         nType("Edm.String"),  _
         tab("Service Terms (Code)"),  _
         Pos(30),  _
         twodBarcode("FLAG")>  _
        Public Property FLAG() As String
            Get
                return _FLAG
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Billable?", value, "^.{0,1}$") then Exit Property
                _IsSetFLAG = True
                If loading Then
                  _FLAG = Value
                Else
                    if not _FLAG = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("FLAG", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _FLAG = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Billing Exempt Type"),  _
         nType("Edm.String"),  _
         tab("Service Terms (Code)"),  _
         Pos(40),  _
         twodBarcode("NOCHARGENAME")>  _
        Public Property NOCHARGENAME() As String
            Get
                return _NOCHARGENAME
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Billing Exempt Type", value, "^.{0,20}$") then Exit Property
                _IsSetNOCHARGENAME = True
                If loading Then
                  _NOCHARGENAME = Value
                Else
                    if not _NOCHARGENAME = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("NOCHARGENAME", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _NOCHARGENAME = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Service Terms (ID)"),  _
         nType("Edm.Int64"),  _
         tab("Service Terms (Code)"),  _
         Pos(10),  _
         Browsable(false),  _
         twodBarcode("SERVTYPE")>  _
        Public Property SERVTYPE() As nullable (of int64)
            Get
                return _SERVTYPE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Service Terms (ID)", value, "^[0-9\-]+$") then Exit Property
                _IsSetSERVTYPE = True
                If loading Then
                  _SERVTYPE = Value
                Else
                    if not _SERVTYPE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("SERVTYPE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _SERVTYPE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        Protected Friend Overrides Sub toJson(ByRef jw As Newtonsoft.Json.JsonTextWriter)
            Dim f as boolean = false
            if _IsSetSERVTCODE then
              if f then
                  jw.WriteRaw(", ""SERVTCODE"": ")
              else
                  jw.WriteRaw("""SERVTCODE"": ")
                  f = true
              end if
              jw.WriteValue(me.SERVTCODE)
            end if
            if _IsSetFLAG then
              if f then
                  jw.WriteRaw(", ""FLAG"": ")
              else
                  jw.WriteRaw("""FLAG"": ")
                  f = true
              end if
              jw.WriteValue(me.FLAG)
            end if
            if _IsSetNOCHARGENAME then
              if f then
                  jw.WriteRaw(", ""NOCHARGENAME"": ")
              else
                  jw.WriteRaw("""NOCHARGENAME"": ")
                  f = true
              end if
              jw.WriteValue(me.NOCHARGENAME)
            end if
            if _IsSetSERVTYPE then
              if f then
                  jw.WriteRaw(", ""SERVTYPE"": ")
              else
                  jw.WriteRaw("""SERVTYPE"": ")
                  f = true
              end if
              jw.WriteValue(me.SERVTYPE)
            end if
        End Sub
        
        Protected Friend Overrides Sub toXML(ByRef xw As System.Xml.XmlWriter, ByVal name As String)
            with xw
              .WriteStartElement("form")
              if name is nothing then
                .WriteAttributeString("name", "FAMILYSERVTYPES")
              else
                .WriteAttributeString("name", name)
              end if
              .WriteAttributeString("result", 0)
              .WriteStartElement("key")
              .WriteAttributeString("name", "SERVTYPE")
              .WriteAttributeString("value", "")
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
            if _IsSetSERVTCODE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "SERVTCODE")
              .WriteAttributeString("value", me.SERVTCODE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "3")
              .WriteEndElement
            end if
            if _IsSetFLAG then
              .WriteStartElement("field")
              .WriteAttributeString("name", "FLAG")
              .WriteAttributeString("value", me.FLAG)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetNOCHARGENAME then
              .WriteStartElement("field")
              .WriteAttributeString("name", "NOCHARGENAME")
              .WriteAttributeString("value", me.NOCHARGENAME)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "20")
              .WriteEndElement
            end if
            if _IsSetSERVTYPE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "SERVTYPE")
              .WriteAttributeString("value", me.SERVTYPE)
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
            end if
              .WriteEndElement
            end with
        End Sub
        
        Protected Friend Overrides Sub HandlesEdit(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILYSERVTYPES = JsonConvert.DeserializeObject(Of FAMILYSERVTYPES)(e.StreamReader.ReadToEnd)
                With obj
                  _SERVTCODE = .SERVTCODE
                  _SERVTDES = .SERVTDES
                  _FLAG = .FLAG
                  _NOCHARGENAME = .NOCHARGENAME
                  _SERVTYPE = .SERVTYPE
                end with
            End If
        End Sub
    End Class
    
    <QueryTitle("Required Analyses")>  _
    Public Class QUERY_FAMILYREQLABANALYSES
        Inherits oDataQuery
        
        Private _Name As String
        
        Private _Value As list(of FAMILYREQLABANALYSES)
        
        Private _Parent As oDataObject
        
        Private _BindingSource As System.Windows.Forms.BindingSource
        
        Public Sub New()
            MyBase.New
            _Value = new list (of FAMILYREQLABANALYSES)
            _Parent = nothing
            _Name = "FAMILYREQLABANALYSES"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            .add(0, "Possible Results")
            end with
        End Sub
        
        Public Sub New(ByRef Parent As oDataObject)
            MyBase.New
            _Value = new list (of FAMILYREQLABANALYSES)
            _Parent = Parent
            _name = "FAMILYREQLABANALYSES_SUBFORM"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            .add(0, "Possible Results")
            end with
        End Sub
        
        <JsonProperty("@odata.context")>  _
        Protected Friend Property Name() As String
            Get
                return _Name
            End Get
            Set
                _Name = value
            End Set
        End Property
        
        Public Property Value() As list(of FAMILYREQLABANALYSES)
            Get
                return _Value
            End Get
            Set
                _Value = value
            End Set
        End Property
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                return _Name
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        Public Overrides ReadOnly Property ObjectType() As System.Type
            Get
                return GetType(FAMILYREQLABANALYSES)
            End Get
        End Property
        
        Public Overrides ReadOnly Property BindingSource() As System.Windows.Forms.BindingSource
            Get
                Return _BindingSource
            End Get
        End Property
        
        Protected Friend Overrides Sub Deserialise(ByRef Stream As StreamReader)
            _value.clear
            For Each _FAMILYREQLABANALYSES As FAMILYREQLABANALYSES In JsonConvert.DeserializeObject(Of QUERY_FAMILYREQLABANALYSES)(stream.ReadToEnd).Value
              With _FAMILYREQLABANALYSES
                  .Parent = Parent
                  .loading = false
              end with
              _value.add(_FAMILYREQLABANALYSES)
            next
            _BindingSource.DataSource = _Value
        End Sub
        
        Protected Friend Overrides Sub HandlesAdd(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILYREQLABANALYSES = JsonConvert.DeserializeObject(Of FAMILYREQLABANALYSES)(e.StreamReader.ReadToEnd)
                With TryCast(BindingSource.Current, FAMILYREQLABANALYSES)
                  .ANALYSISCODE = obj.ANALYSISCODE
                  .ANALYSISDES = obj.ANALYSISDES
                  .METHODNAME = obj.METHODNAME
                  .MINVALUE = obj.MINVALUE
                  .MAXVALUE = obj.MAXVALUE
                  .PRINTFLAG = obj.PRINTFLAG
                  .ANALYSIS = obj.ANALYSIS
                  .KEY2 = obj.KEY2
                  .REQANALYSIS = obj.REQANALYSIS
                  .TYPE = obj.TYPE
                  .loading = false
                end with
            End If
        End Sub
        
        Private Sub HandlesBeginAdd(ByVal sender As Object, ByVal e As ComponentModel.AddingNewEventArgs)
            e.NewObject = new FAMILYREQLABANALYSES(me.Parent)
        End Sub
        
        Protected Friend Overrides Sub Remove(ByRef obj As oDataObject)
            for each _FAMILYREQLABANALYSES as FAMILYREQLABANALYSES in value
              If _FAMILYREQLABANALYSES.Equals(trycast(obj,FAMILYREQLABANALYSES)) Then
                  value.remove(_FAMILYREQLABANALYSES)
                  exit for
              end if
            next
        End Sub
    End Class
    
    Public Class FAMILYREQLABANALYSES
        Inherits oDataObject
        
        Private _Parent As oDataObject
        
        Private _IsSetANALYSISCODE As Boolean = Boolean.FalseString
        
        Private _ANALYSISCODE As String
        
        Private _ANALYSISDES As String
        
        Private _IsSetMETHODNAME As Boolean = Boolean.FalseString
        
        Private _METHODNAME As String
        
        Private _IsSetMINVALUE As Boolean = Boolean.FalseString
        
        Private _MINVALUE As String
        
        Private _IsSetMAXVALUE As Boolean = Boolean.FalseString
        
        Private _MAXVALUE As String
        
        Private _IsSetPRINTFLAG As Boolean = Boolean.FalseString
        
        Private _PRINTFLAG As String
        
        Private _IsSetANALYSIS As Boolean = Boolean.FalseString
        
        Private _ANALYSIS As Long
        
        Private _IsSetKEY2 As Boolean = Boolean.FalseString
        
        Private _KEY2 As Long
        
        Private _IsSetREQANALYSIS As Boolean = Boolean.FalseString
        
        Private _REQANALYSIS As Long
        
        Private _IsSetTYPE As Boolean = Boolean.FalseString
        
        Private _TYPE As String
        
        Private _LABANALYSESRESULTS_SUBFORM As QUERY_LABANALYSESRESULTS
        
        Public Sub New(ByVal Parent As oDataObject)
            MyBase.New
            _parent = Parent
            ChildQuery.add(0, new oNavigation("Possible Results"))
            _LABANALYSESRESULTS_SUBFORM = new QUERY_LABANALYSESRESULTS(me)
            WITH ChildQuery(0)
               .setoDataQuery(_LABANALYSESRESULTS_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Possible Results", _LABANALYSESRESULTS_SUBFORM))
               end with
            end with
        End Sub
        
        Public Sub New()
            MyBase.New
            ChildQuery.add(0, new oNavigation("Possible Results"))
            _LABANALYSESRESULTS_SUBFORM = new QUERY_LABANALYSESRESULTS(me)
            WITH ChildQuery(0)
               .setoDataQuery(_LABANALYSESRESULTS_SUBFORM)
               WITH .oDataQuery.SibligQuery
                   .add(0, new oNavigation("Possible Results", _LABANALYSESRESULTS_SUBFORM))
               end with
            end with
        End Sub
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                if _parent is nothing then
                    return "FAMILYREQLABANALYSES"
                else
                    return "FAMILYREQLABANALYSES_SUBFORM"
                end if
            End Get
        End Property
        
        Public Overrides ReadOnly Property KeyString() As String
            Get
                return string.format( _
                 "ANALYSIS={0},KEY2={1},TYPE={2}", _
                  string.format("{0}",ANALYSIS), _
                  string.format("{0}",KEY2), _
                  string.format("'{0}'",TYPE) _ 
               )
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        <DisplayName("Analysis Code"),  _
         nType("Edm.String"),  _
         tab("Analysis Code"),  _
         Pos(10),  _
         Mandatory(true),  _
         twodBarcode("ANALYSISCODE")>  _
        Public Property ANALYSISCODE() As String
            Get
                return _ANALYSISCODE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Analysis Code", value, "^.{0,10}$") then Exit Property
                _IsSetANALYSISCODE = True
                If loading Then
                  _ANALYSISCODE = Value
                Else
                    if not _ANALYSISCODE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("ANALYSISCODE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _ANALYSISCODE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Analysis Description"),  _
         nType("Edm.String"),  _
         tab("Analysis Code"),  _
         Pos(15),  _
         [ReadOnly](true),  _
         twodBarcode("ANALYSISDES")>  _
        Public Property ANALYSISDES() As String
            Get
                return _ANALYSISDES
            End Get
            Set
                if not(value is nothing) then
                  _ANALYSISDES = Value
                end if
            End Set
        End Property
        
        <DisplayName("Method of Analysis"),  _
         nType("Edm.String"),  _
         tab("Analysis Code"),  _
         Pos(18),  _
         twodBarcode("METHODNAME")>  _
        Public Property METHODNAME() As String
            Get
                return _METHODNAME
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Method of Analysis", value, "^.{0,20}$") then Exit Property
                _IsSetMETHODNAME = True
                If loading Then
                  _METHODNAME = Value
                Else
                    if not _METHODNAME = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("METHODNAME", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _METHODNAME = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Minimum Value"),  _
         nType("Edm.String"),  _
         tab("Analysis Code"),  _
         Pos(20),  _
         twodBarcode("MINVALUE")>  _
        Public Property MINVALUE() As String
            Get
                return _MINVALUE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Minimum Value", value, "^.{0,24}$") then Exit Property
                _IsSetMINVALUE = True
                If loading Then
                  _MINVALUE = Value
                Else
                    if not _MINVALUE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("MINVALUE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _MINVALUE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Maximum Value"),  _
         nType("Edm.String"),  _
         tab("Analysis Code"),  _
         Pos(30),  _
         twodBarcode("MAXVALUE")>  _
        Public Property MAXVALUE() As String
            Get
                return _MAXVALUE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Maximum Value", value, "^.{0,24}$") then Exit Property
                _IsSetMAXVALUE = True
                If loading Then
                  _MAXVALUE = Value
                Else
                    if not _MAXVALUE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("MAXVALUE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _MAXVALUE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Display in Printout?"),  _
         nType("Edm.String"),  _
         tab("Analysis Code"),  _
         Pos(40),  _
         twodBarcode("PRINTFLAG")>  _
        Public Property PRINTFLAG() As String
            Get
                return _PRINTFLAG
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Display in Printout?", value, "^.{0,1}$") then Exit Property
                _IsSetPRINTFLAG = True
                If loading Then
                  _PRINTFLAG = Value
                Else
                    if not _PRINTFLAG = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("PRINTFLAG", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _PRINTFLAG = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Analysis (ID)"),  _
         nType("Edm.Int64"),  _
         tab("Analysis Code"),  _
         Pos(99),  _
         Browsable(false),  _
         twodBarcode("ANALYSIS")>  _
        Public Property ANALYSIS() As nullable (of int64)
            Get
                return _ANALYSIS
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Analysis (ID)", value, "^[0-9\-]+$") then Exit Property
                _IsSetANALYSIS = True
                If loading Then
                  _ANALYSIS = Value
                Else
                    if not _ANALYSIS = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("ANALYSIS", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _ANALYSIS = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("KEY2"),  _
         nType("Edm.Int64"),  _
         tab("Analysis Code"),  _
         Pos(0),  _
         twodBarcode("KEY2")>  _
        Public Property KEY2() As nullable (of int64)
            Get
                return _KEY2
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("KEY2", value, "^[0-9\-]+$") then Exit Property
                _IsSetKEY2 = True
                If loading Then
                  _KEY2 = Value
                Else
                    if not _KEY2 = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("KEY2", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _KEY2 = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Required Analysis-ID"),  _
         nType("Edm.Int64"),  _
         tab("Required Analysis-ID"),  _
         Pos(99),  _
         Browsable(false),  _
         twodBarcode("REQANALYSIS")>  _
        Public Property REQANALYSIS() As nullable (of int64)
            Get
                return _REQANALYSIS
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Required Analysis-ID", value, "^[0-9\-]+$") then Exit Property
                _IsSetREQANALYSIS = True
                If loading Then
                  _REQANALYSIS = Value
                Else
                    if not _REQANALYSIS = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("REQANALYSIS", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _REQANALYSIS = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Type"),  _
         nType("Edm.String"),  _
         tab("Required Analysis-ID"),  _
         Pos(99),  _
         Browsable(false),  _
         twodBarcode("TYPE")>  _
        Public Property TYPE() As String
            Get
                return _TYPE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Type", value, "^.{0,8}$") then Exit Property
                _IsSetTYPE = True
                If loading Then
                  _TYPE = Value
                Else
                    if not _TYPE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("TYPE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _TYPE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <JsonIgnore(),  _
         Browsable(false)>  _
        Public Property LABANALYSESRESULTS_SUBFORM() As QUERY_LABANALYSESRESULTS
            Get
                return _LABANALYSESRESULTS_SUBFORM
            End Get
            Set
                _LABANALYSESRESULTS_SUBFORM = value
            End Set
        End Property
        
        Protected Friend Overrides Sub toJson(ByRef jw As Newtonsoft.Json.JsonTextWriter)
            Dim f as boolean = false
            if _IsSetANALYSISCODE then
              if f then
                  jw.WriteRaw(", ""ANALYSISCODE"": ")
              else
                  jw.WriteRaw("""ANALYSISCODE"": ")
                  f = true
              end if
              jw.WriteValue(me.ANALYSISCODE)
            end if
            if _IsSetMETHODNAME then
              if f then
                  jw.WriteRaw(", ""METHODNAME"": ")
              else
                  jw.WriteRaw("""METHODNAME"": ")
                  f = true
              end if
              jw.WriteValue(me.METHODNAME)
            end if
            if _IsSetMINVALUE then
              if f then
                  jw.WriteRaw(", ""MINVALUE"": ")
              else
                  jw.WriteRaw("""MINVALUE"": ")
                  f = true
              end if
              jw.WriteValue(me.MINVALUE)
            end if
            if _IsSetMAXVALUE then
              if f then
                  jw.WriteRaw(", ""MAXVALUE"": ")
              else
                  jw.WriteRaw("""MAXVALUE"": ")
                  f = true
              end if
              jw.WriteValue(me.MAXVALUE)
            end if
            if _IsSetPRINTFLAG then
              if f then
                  jw.WriteRaw(", ""PRINTFLAG"": ")
              else
                  jw.WriteRaw("""PRINTFLAG"": ")
                  f = true
              end if
              jw.WriteValue(me.PRINTFLAG)
            end if
            if _IsSetANALYSIS then
              if f then
                  jw.WriteRaw(", ""ANALYSIS"": ")
              else
                  jw.WriteRaw("""ANALYSIS"": ")
                  f = true
              end if
              jw.WriteValue(me.ANALYSIS)
            end if
            if _IsSetKEY2 then
              if f then
                  jw.WriteRaw(", ""KEY2"": ")
              else
                  jw.WriteRaw("""KEY2"": ")
                  f = true
              end if
              jw.WriteValue(me.KEY2)
            end if
            if _IsSetREQANALYSIS then
              if f then
                  jw.WriteRaw(", ""REQANALYSIS"": ")
              else
                  jw.WriteRaw("""REQANALYSIS"": ")
                  f = true
              end if
              jw.WriteValue(me.REQANALYSIS)
            end if
            if _IsSetTYPE then
              if f then
                  jw.WriteRaw(", ""TYPE"": ")
              else
                  jw.WriteRaw("""TYPE"": ")
                  f = true
              end if
              jw.WriteValue(me.TYPE)
            end if
            if _LABANALYSESRESULTS_SUBFORM.value.count > 0 then
              jw.WriteRaw(", LABANALYSESRESULTS_SUBFORM: ")
              jw.WriteRaw("[ ")
              dim i as integer = 0
              for each itm as LABANALYSESRESULTS in _LABANALYSESRESULTS_SUBFORM.value
                jw.WriteRaw("{ ")
                itm.tojson(jw)
                jw.WriteRaw(" }")
                if i < _LABANALYSESRESULTS_SUBFORM.value.count - 1 then jw.WriteRaw(", ")
                i += 1
                next
              jw.WriteRaw(" ]")
            end if
        End Sub
        
        Protected Friend Overrides Sub toXML(ByRef xw As System.Xml.XmlWriter, ByVal name As String)
            with xw
              .WriteStartElement("form")
              if name is nothing then
                .WriteAttributeString("name", "FAMILYREQLABANALYSES")
              else
                .WriteAttributeString("name", name)
              end if
              .WriteAttributeString("result", 0)
              .WriteStartElement("key")
              .WriteAttributeString("name", "ANALYSIS")
              .WriteAttributeString("value", "")
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
              .WriteStartElement("key")
              .WriteAttributeString("name", "KEY2")
              .WriteAttributeString("value", "")
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
              .WriteStartElement("key")
              .WriteAttributeString("name", "TYPE")
              .WriteAttributeString("value", "")
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "8")
              .WriteEndElement
            if _IsSetANALYSISCODE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "ANALYSISCODE")
              .WriteAttributeString("value", me.ANALYSISCODE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "10")
              .WriteEndElement
            end if
            if _IsSetMETHODNAME then
              .WriteStartElement("field")
              .WriteAttributeString("name", "METHODNAME")
              .WriteAttributeString("value", me.METHODNAME)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "20")
              .WriteEndElement
            end if
            if _IsSetMINVALUE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "MINVALUE")
              .WriteAttributeString("value", me.MINVALUE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "24")
              .WriteEndElement
            end if
            if _IsSetMAXVALUE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "MAXVALUE")
              .WriteAttributeString("value", me.MAXVALUE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "24")
              .WriteEndElement
            end if
            if _IsSetPRINTFLAG then
              .WriteStartElement("field")
              .WriteAttributeString("name", "PRINTFLAG")
              .WriteAttributeString("value", me.PRINTFLAG)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetANALYSIS then
              .WriteStartElement("field")
              .WriteAttributeString("name", "ANALYSIS")
              .WriteAttributeString("value", me.ANALYSIS)
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
            end if
            if _IsSetKEY2 then
              .WriteStartElement("field")
              .WriteAttributeString("name", "KEY2")
              .WriteAttributeString("value", me.KEY2)
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
            end if
            if _IsSetREQANALYSIS then
              .WriteStartElement("field")
              .WriteAttributeString("name", "REQANALYSIS")
              .WriteAttributeString("value", me.REQANALYSIS)
              .WriteAttributeString("type", "Edm.Int64")
              .WriteEndElement
            end if
            if _IsSetTYPE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "TYPE")
              .WriteAttributeString("value", me.TYPE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "8")
              .WriteEndElement
            end if
            if _LABANALYSESRESULTS_SUBFORM.value.count > 0 then
              for each itm as LABANALYSESRESULTS in _LABANALYSESRESULTS_SUBFORM.Value
                itm.toXML(xw,"LABANALYSESRESULTS_SUBFORM")
              next
            end if
              .WriteEndElement
            end with
        End Sub
        
        Protected Friend Overrides Sub HandlesEdit(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as FAMILYREQLABANALYSES = JsonConvert.DeserializeObject(Of FAMILYREQLABANALYSES)(e.StreamReader.ReadToEnd)
                With obj
                  _ANALYSISCODE = .ANALYSISCODE
                  _ANALYSISDES = .ANALYSISDES
                  _METHODNAME = .METHODNAME
                  _MINVALUE = .MINVALUE
                  _MAXVALUE = .MAXVALUE
                  _PRINTFLAG = .PRINTFLAG
                  _ANALYSIS = .ANALYSIS
                  _KEY2 = .KEY2
                  _REQANALYSIS = .REQANALYSIS
                  _TYPE = .TYPE
                end with
            End If
        End Sub
    End Class
    
    Public Enum Subform_FAMILYREQLABANALYSES
        
        LABANALYSESRESULTS = 0
    End Enum
    
    <QueryTitle("Possible Results")>  _
    Public Class QUERY_LABANALYSESRESULTS
        Inherits oDataQuery
        
        Private _Name As String
        
        Private _Value As list(of LABANALYSESRESULTS)
        
        Private _Parent As oDataObject
        
        Private _BindingSource As System.Windows.Forms.BindingSource
        
        Public Sub New()
            MyBase.New
            _Value = new list (of LABANALYSESRESULTS)
            _Parent = nothing
            _Name = "LABANALYSESRESULTS"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            end with
        End Sub
        
        Public Sub New(ByRef Parent As oDataObject)
            MyBase.New
            _Value = new list (of LABANALYSESRESULTS)
            _Parent = Parent
            _name = "LABANALYSESRESULTS_SUBFORM"
            _BindingSource = new BindingSource
            AddHandler _BindingSource.AddingNew, AddressOf HandlesBeginAdd
            with ChildQuery
            end with
        End Sub
        
        <JsonProperty("@odata.context")>  _
        Protected Friend Property Name() As String
            Get
                return _Name
            End Get
            Set
                _Name = value
            End Set
        End Property
        
        Public Property Value() As list(of LABANALYSESRESULTS)
            Get
                return _Value
            End Get
            Set
                _Value = value
            End Set
        End Property
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                return _Name
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        Public Overrides ReadOnly Property ObjectType() As System.Type
            Get
                return GetType(LABANALYSESRESULTS)
            End Get
        End Property
        
        Public Overrides ReadOnly Property BindingSource() As System.Windows.Forms.BindingSource
            Get
                Return _BindingSource
            End Get
        End Property
        
        Protected Friend Overrides Sub Deserialise(ByRef Stream As StreamReader)
            _value.clear
            For Each _LABANALYSESRESULTS As LABANALYSESRESULTS In JsonConvert.DeserializeObject(Of QUERY_LABANALYSESRESULTS)(stream.ReadToEnd).Value
              With _LABANALYSESRESULTS
                  .Parent = Parent
                  .loading = false
              end with
              _value.add(_LABANALYSESRESULTS)
            next
            _BindingSource.DataSource = _Value
        End Sub
        
        Protected Friend Overrides Sub HandlesAdd(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as LABANALYSESRESULTS = JsonConvert.DeserializeObject(Of LABANALYSESRESULTS)(e.StreamReader.ReadToEnd)
                With TryCast(BindingSource.Current, LABANALYSESRESULTS)
                  .RESCODE = obj.RESCODE
                  .VALID = obj.VALID
                  .NOTVALID = obj.NOTVALID
                  .QRANKCODE = obj.QRANKCODE
                  .FROMVAL = obj.FROMVAL
                  .TOVAL = obj.TOVAL
                  .loading = false
                end with
            End If
        End Sub
        
        Private Sub HandlesBeginAdd(ByVal sender As Object, ByVal e As ComponentModel.AddingNewEventArgs)
            e.NewObject = new LABANALYSESRESULTS(me.Parent)
        End Sub
        
        Protected Friend Overrides Sub Remove(ByRef obj As oDataObject)
            for each _LABANALYSESRESULTS as LABANALYSESRESULTS in value
              If _LABANALYSESRESULTS.Equals(trycast(obj,LABANALYSESRESULTS)) Then
                  value.remove(_LABANALYSESRESULTS)
                  exit for
              end if
            next
        End Sub
    End Class
    
    Public Class LABANALYSESRESULTS
        Inherits oDataObject
        
        Private _Parent As oDataObject
        
        Private _IsSetRESCODE As Boolean = Boolean.FalseString
        
        Private _RESCODE As String
        
        Private _IsSetVALID As Boolean = Boolean.FalseString
        
        Private _VALID As String
        
        Private _IsSetNOTVALID As Boolean = Boolean.FalseString
        
        Private _NOTVALID As String
        
        Private _IsSetQRANKCODE As Boolean = Boolean.FalseString
        
        Private _QRANKCODE As String
        
        Private _IsSetFROMVAL As Boolean = Boolean.FalseString
        
        Private _FROMVAL As String
        
        Private _IsSetTOVAL As Boolean = Boolean.FalseString
        
        Private _TOVAL As String
        
        Public Sub New(ByVal Parent As oDataObject)
            MyBase.New
            _parent = Parent
        End Sub
        
        Public Sub New()
            MyBase.New
        End Sub
        
        Protected Friend Overrides ReadOnly Property EntityName() As String
            Get
                if _parent is nothing then
                    return "LABANALYSESRESULTS"
                else
                    return "LABANALYSESRESULTS_SUBFORM"
                end if
            End Get
        End Property
        
        Public Overrides ReadOnly Property KeyString() As String
            Get
                return string.format( _
                 "RESCODE={0}", _
                  string.format("'{0}'",RESCODE) _ 
               )
            End Get
        End Property
        
        Public Overrides Property Parent() As oDataObject
            Get
                return _Parent
            End Get
            Set
                _Parent = value
            End Set
        End Property
        
        <DisplayName("Possible Result"),  _
         nType("Edm.String"),  _
         tab("Possible Result"),  _
         Pos(10),  _
         Mandatory(true),  _
         twodBarcode("RESCODE")>  _
        Public Property RESCODE() As String
            Get
                return _RESCODE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Possible Result", value, "^.{0,24}$") then Exit Property
                _IsSetRESCODE = True
                If loading Then
                  _RESCODE = Value
                Else
                    if not _RESCODE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("RESCODE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _RESCODE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Acceptable?"),  _
         nType("Edm.String"),  _
         tab("Possible Result"),  _
         Pos(20),  _
         twodBarcode("VALID")>  _
        Public Property VALID() As String
            Get
                return _VALID
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Acceptable?", value, "^.{0,1}$") then Exit Property
                _IsSetVALID = True
                If loading Then
                  _VALID = Value
                Else
                    if not _VALID = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("VALID", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _VALID = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Unacceptable?"),  _
         nType("Edm.String"),  _
         tab("Possible Result"),  _
         Pos(30),  _
         twodBarcode("NOTVALID")>  _
        Public Property NOTVALID() As String
            Get
                return _NOTVALID
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Unacceptable?", value, "^.{0,1}$") then Exit Property
                _IsSetNOTVALID = True
                If loading Then
                  _NOTVALID = Value
                Else
                    if not _NOTVALID = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("NOTVALID", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _NOTVALID = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Quality Code"),  _
         nType("Edm.String"),  _
         tab("Possible Result"),  _
         Pos(40),  _
         twodBarcode("QRANKCODE")>  _
        Public Property QRANKCODE() As String
            Get
                return _QRANKCODE
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Quality Code", value, "^.{0,4}$") then Exit Property
                _IsSetQRANKCODE = True
                If loading Then
                  _QRANKCODE = Value
                Else
                    if not _QRANKCODE = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("QRANKCODE", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _QRANKCODE = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Minimum QualityValue"),  _
         nType("Edm.String"),  _
         tab("Possible Result"),  _
         Pos(50),  _
         twodBarcode("FROMVAL")>  _
        Public Property FROMVAL() As String
            Get
                return _FROMVAL
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Minimum QualityValue", value, "^.{0,24}$") then Exit Property
                _IsSetFROMVAL = True
                If loading Then
                  _FROMVAL = Value
                Else
                    if not _FROMVAL = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("FROMVAL", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _FROMVAL = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        <DisplayName("Maximum QualityValue"),  _
         nType("Edm.String"),  _
         tab("Possible Result"),  _
         Pos(60),  _
         twodBarcode("TOVAL")>  _
        Public Property TOVAL() As String
            Get
                return _TOVAL
            End Get
            Set
                if value is nothing then Exit Property
                if not mybase.validate("Maximum QualityValue", value, "^.{0,24}$") then Exit Property
                _IsSetTOVAL = True
                If loading Then
                  _TOVAL = Value
                Else
                    if not _TOVAL = value then
                        Connection.RaiseStartData()
                        loading = true
                        Dim cn As New oDataPUT(Me, PropertyStream("TOVAL", Value), AddressOf HandlesEdit)
                        loading = false
                        If Connection.LastError is nothing Then
                            _TOVAL = Value
                        End If
                    end if
                end if
            End Set
        End Property
        
        Protected Friend Overrides Sub toJson(ByRef jw As Newtonsoft.Json.JsonTextWriter)
            Dim f as boolean = false
            if _IsSetRESCODE then
              if f then
                  jw.WriteRaw(", ""RESCODE"": ")
              else
                  jw.WriteRaw("""RESCODE"": ")
                  f = true
              end if
              jw.WriteValue(me.RESCODE)
            end if
            if _IsSetVALID then
              if f then
                  jw.WriteRaw(", ""VALID"": ")
              else
                  jw.WriteRaw("""VALID"": ")
                  f = true
              end if
              jw.WriteValue(me.VALID)
            end if
            if _IsSetNOTVALID then
              if f then
                  jw.WriteRaw(", ""NOTVALID"": ")
              else
                  jw.WriteRaw("""NOTVALID"": ")
                  f = true
              end if
              jw.WriteValue(me.NOTVALID)
            end if
            if _IsSetQRANKCODE then
              if f then
                  jw.WriteRaw(", ""QRANKCODE"": ")
              else
                  jw.WriteRaw("""QRANKCODE"": ")
                  f = true
              end if
              jw.WriteValue(me.QRANKCODE)
            end if
            if _IsSetFROMVAL then
              if f then
                  jw.WriteRaw(", ""FROMVAL"": ")
              else
                  jw.WriteRaw("""FROMVAL"": ")
                  f = true
              end if
              jw.WriteValue(me.FROMVAL)
            end if
            if _IsSetTOVAL then
              if f then
                  jw.WriteRaw(", ""TOVAL"": ")
              else
                  jw.WriteRaw("""TOVAL"": ")
                  f = true
              end if
              jw.WriteValue(me.TOVAL)
            end if
        End Sub
        
        Protected Friend Overrides Sub toXML(ByRef xw As System.Xml.XmlWriter, ByVal name As String)
            with xw
              .WriteStartElement("form")
              if name is nothing then
                .WriteAttributeString("name", "LABANALYSESRESULTS")
              else
                .WriteAttributeString("name", name)
              end if
              .WriteAttributeString("result", 0)
              .WriteStartElement("key")
              .WriteAttributeString("name", "RESCODE")
              .WriteAttributeString("value", "")
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "24")
              .WriteEndElement
            if _IsSetRESCODE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "RESCODE")
              .WriteAttributeString("value", me.RESCODE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "24")
              .WriteEndElement
            end if
            if _IsSetVALID then
              .WriteStartElement("field")
              .WriteAttributeString("name", "VALID")
              .WriteAttributeString("value", me.VALID)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetNOTVALID then
              .WriteStartElement("field")
              .WriteAttributeString("name", "NOTVALID")
              .WriteAttributeString("value", me.NOTVALID)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "1")
              .WriteEndElement
            end if
            if _IsSetQRANKCODE then
              .WriteStartElement("field")
              .WriteAttributeString("name", "QRANKCODE")
              .WriteAttributeString("value", me.QRANKCODE)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "4")
              .WriteEndElement
            end if
            if _IsSetFROMVAL then
              .WriteStartElement("field")
              .WriteAttributeString("name", "FROMVAL")
              .WriteAttributeString("value", me.FROMVAL)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "24")
              .WriteEndElement
            end if
            if _IsSetTOVAL then
              .WriteStartElement("field")
              .WriteAttributeString("name", "TOVAL")
              .WriteAttributeString("value", me.TOVAL)
              .WriteAttributeString("type", "Edm.String")
              .WriteAttributeString("MaxLength", "24")
              .WriteEndElement
            end if
              .WriteEndElement
            end with
        End Sub
        
        Protected Friend Overrides Sub HandlesEdit(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException is nothing Then
                Connection.LastError = e.InterfaceException
            Else
                dim obj as LABANALYSESRESULTS = JsonConvert.DeserializeObject(Of LABANALYSESRESULTS)(e.StreamReader.ReadToEnd)
                With obj
                  _RESCODE = .RESCODE
                  _VALID = .VALID
                  _NOTVALID = .NOTVALID
                  _QRANKCODE = .QRANKCODE
                  _FROMVAL = .FROMVAL
                  _TOVAL = .TOVAL
                end with
            End If
        End Sub
    End Class
End Namespace
