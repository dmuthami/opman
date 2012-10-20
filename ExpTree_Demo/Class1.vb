
'wrapping up forms
Public Class myForms

#Region "public members"

#End Region

#Region "codes rahisi rahisi"
    Private Shared m_CustomerForm As frmAddLead
    Public Shared Property CustomerForm() As frmAddLead
        Get
            Return m_CustomerForm
        End Get
        Set(ByVal Value As frmAddLead)
            m_CustomerForm = Value
        End Set
    End Property
    Private Shared m_CustomerForm1 As frmEditcontact
    Public Shared Property CustomerForm1() As frmEditcontact
        Get
            Return m_CustomerForm1
        End Get
        Set(ByVal Value As frmEditcontact)
            m_CustomerForm1 = Value
        End Set
    End Property
    Private Shared m_CustomerForm2 As frmEditJob
    Public Shared Property CustomerForm2() As frmEditJob
        Get
            Return m_CustomerForm2
        End Get
        Set(ByVal Value As frmEditJob)
            m_CustomerForm2 = Value
        End Set
    End Property
    Private Shared m_CustomerForm3 As frmMe
    Public Shared Property CustomerForm3() As frmMe
        Get
            Return m_CustomerForm3
        End Get
        Set(ByVal Value As frmMe)
            m_CustomerForm3 = Value
        End Set
    End Property
    Private Shared m_CustomerForm4 As frmEditLead
    Public Shared Property CustomerForm4() As frmEditLead
        Get
            Return m_CustomerForm4
        End Get
        Set(ByVal Value As frmEditLead)
            m_CustomerForm4 = Value
        End Set
    End Property
    Private Shared m_uid As String
    Public Shared Property id_no() As String
        Get
            Return m_uid
        End Get
        Set(ByVal Value As String)
            m_uid = Value
        End Set
    End Property
    'Public Shared Property pwd() As String
    '    Get
    '        Return pwd
    '    End Get
    '    Set(ByVal Value As String)
    '        pwd = Value
    '    End Set
    'End Property
    Private Shared m_timesheet As frmtime
    Public Shared Property timesheet() As frmtime
        Get
            Return m_timesheet
        End Get
        Set(ByVal Value As frmtime)
            m_timesheet = Value
        End Set
    End Property
    Private Shared m_npersonnel As frmnewpersonnel
    Public Shared Property npersonnel() As frmnewpersonnel
        Get
            Return m_npersonnel
        End Get
        Set(ByVal Value As frmnewpersonnel)
            m_npersonnel = Value
        End Set
    End Property
    Private Shared m_adminform As frmpersonneladmin
    Public Shared Property adminform() As frmpersonneladmin
        Get
            Return m_adminform
        End Get
        Set(ByVal Value As frmpersonneladmin)
            m_adminform = Value
        End Set
    End Property
    Private Shared m_jobsheet As frmjobsheet
    Public Shared Property jobsheet() As frmjobsheet
        Get
            Return m_jobsheet
        End Get
        Set(ByVal Value As frmjobsheet)
            m_jobsheet = Value
        End Set
    End Property

    Private Shared m_main As frmHome
    Public Shared Property Main() As frmHome
        Get
            Return m_main
        End Get
        Set(ByVal Value As frmHome)
            m_main = Value
        End Set
    End Property

    Private Shared m_admin As frmadmin
    Public Shared Property admin() As frmadmin
        Get
            Return m_admin
        End Get
        Set(ByVal Value As frmadmin)
            m_admin = Value
        End Set
    End Property
#End Region

#Region " equipments"
    Public Shared iseditequip As Boolean = False
    Private Shared m_equip As frminventories
    Public Shared Property equipments() As frminventories
        Get
            Return m_equip
        End Get
        Set(ByVal Value As frminventories)
            m_equip = Value
        End Set
    End Property

    Private Shared m_equipedit As frmeditequip
    Public Shared Property editequipments() As frmeditequip
        Get
            Return m_equipedit
        End Get
        Set(ByVal Value As frmeditequip)
            m_equipedit = Value
        End Set
    End Property
    Public Shared isequipactions As Boolean = False
    Private Shared m_equipactions As frmequipmentactions
    Public Shared Property equipactions() As frmequipmentactions
        Get
            Return m_equipactions
        End Get
        Set(ByVal Value As frmequipmentactions)
            m_equipactions = Value
        End Set
    End Property
#End Region

#Region " tojobs"
    Private Shared m_tojobs As frmtojobs
    Public Shared Property tojobs() As frmtojobs
        Get
            Return m_tojobs
        End Get
        Set(ByVal Value As frmtojobs)
            m_tojobs = Value
        End Set
    End Property
    Public Shared iseditassignequip As Boolean = False
    Private Shared m_editassignequip As frmeditequipassign
    Public Shared Property editassignequip() As frmeditequipassign
        Get
            Return m_editassignequip
        End Get
        Set(ByVal Value As frmeditequipassign)
            m_editassignequip = Value
        End Set
    End Property
#End Region

#Region "history"
    Public Shared ishistory As Boolean = False
    Private Shared m_history As frmhistory
    Public Shared Property historry() As frmhistory
        Get
            Return m_history
        End Get
        Set(ByVal Value As frmhistory)
            m_history = Value
        End Set
    End Property
#End Region

#Region "maintenace"
    Public Shared ismaintenace As Boolean = False
    Private Shared m_maintenace As frmservice
    Public Shared Property maintenace() As frmservice
        Get
            Return m_maintenace
        End Get
        Set(ByVal Value As frmservice)
            m_maintenace = Value
        End Set
    End Property
    Public Shared isadddmaintenace As Boolean = False
    Private Shared m_addmaintenace As frmaddmaintenanceinfo
    Public Shared Property addmaintenace() As frmaddmaintenanceinfo
        Get
            Return m_addmaintenace
        End Get
        Set(ByVal Value As frmaddmaintenanceinfo)
            m_addmaintenace = Value
        End Set
    End Property
    Public Shared iseditmaintenace As Boolean = False
    Private Shared m_editmaintenace As frmeditmaintenanceinfo
    Public Shared Property editmaintenace() As frmeditmaintenanceinfo
        Get
            Return m_editmaintenace
        End Get
        Set(ByVal Value As frmeditmaintenanceinfo)
            m_editmaintenace = Value
        End Set
    End Property
#End Region

#Region "job sheet"
    Public Shared issickleave As Boolean = False
    Private Shared m_sickleave As frmaddsickleave
    Public Shared Property sickleave() As frmaddsickleave
        Get
            Return m_sickleave
        End Get
        Set(ByVal Value As frmaddsickleave)
            m_sickleave = Value
        End Set
    End Property

    Public Shared istimeoff As Boolean = False
    Private Shared m_timeoff As frmtimeoff
    Public Shared Property timeoff() As frmtimeoff
        Get
            Return m_timeoff
        End Get
        Set(ByVal Value As frmtimeoff)
            m_timeoff = Value
        End Set
    End Property
#End Region

#Region "jobs"
    Private Shared m_job As frmEditJob
    Public Shared Property jobdetails() As frmEditJob
        Get
            Return m_job
        End Get
        Set(ByVal Value As frmEditJob)
            m_job = Value
        End Set
    End Property

#Region "casuals/accomodation/travel"
    Private Shared m_casuals As frmaddcasual
    Public Shared Property casuals() As frmaddcasual
        Get
            Return m_casuals
        End Get
        Set(ByVal Value As frmaddcasual)
            m_casuals = Value
        End Set
    End Property

    Private Shared m_accomodation As frmaddaccomodation
    Public Shared Property accomodation() As frmaddaccomodation
        Get
            Return m_accomodation
        End Get
        Set(ByVal Value As frmaddaccomodation)
            m_accomodation = Value
        End Set
    End Property
    Private Shared m_travel As frmtravel
    Public Shared Property travel() As frmtravel
        Get
            Return m_travel
        End Get
        Set(ByVal Value As frmtravel)
            m_travel = Value
        End Set
    End Property
#End Region

#Region "hired/ramani"
    Private Shared m_hired As frmhiredequip
    Public Shared Property hired() As frmhiredequip
        Get
            Return m_hired
        End Get
        Set(ByVal Value As frmhiredequip)
            m_hired = Value
        End Set
    End Property
#End Region

#End Region

#Region "it issues"
    Private Shared m_it As frmit
    Public Shared Property itissues() As frmit
        Get
            Return m_it
        End Get
        Set(ByVal Value As frmit)
            m_it = Value
        End Set
    End Property
#End Region

#Region "query jobs"
    Private Shared m_qjobs As frmjobs
    Public Shared Property qjobs() As frmjobs
        Get
            Return m_qjobs
        End Get
        Set(ByVal Value As frmjobs)
            m_qjobs = Value
        End Set
    End Property
#End Region

#Region "connection settings"
    Private Shared m_qconnstr As String
    Public Shared Property qconnstr() As String
        Get
            Return m_qconnstr
        End Get
        Set(ByVal Value As String)
            m_qconnstr = Value
        End Set
    End Property
    'Private Shared m_arrcon() As String
    'Public Shared Property arrcon()
    '    Get
    '        Return m_arrcon
    '    End Get
    '    Set(ByVal Value)
    '        m_arrcon = Value
    '    End Set
    'End Property
    Private Shared m_qfolderpath As String
    Public Shared Property qfolderpath() As String
        Get
            Return m_qfolderpath
        End Get
        Set(ByVal Value As String)
            m_qfolderpath = Value
        End Set
    End Property
    Private Shared m_str_r As String
    Public Shared Property str_r() As String
        Get
            Return m_str_r
        End Get
        Set(ByVal Value As String)
            m_str_r = Value
        End Set
    End Property
    Private Shared m_mailserver As String
    Public Shared Property mailserver() As String
        Get
            Return m_mailserver
        End Get
        Set(ByVal Value As String)
            m_mailserver = Value
        End Set
    End Property
#End Region


End Class

