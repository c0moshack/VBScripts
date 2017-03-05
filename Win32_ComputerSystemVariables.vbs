Sub displaycomputersystem
	' Connect to WMI for Computer System
	 strComputer = "."
	 Set objWMIService = GetObject("winmgmts:\\" & strComputer &  "\root\cimv2")
	 Set colComputerSystems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")

	 ' Get WMI Variables
	 For Each objComputerSystem in colComputerSystems
		'Get Computer Name
		Wscript.Echo "AdminPasswordStatus = " & objComputerSystem.AdminPasswordStatus
		Wscript.Echo "AutomaticManagedPagefile = " & objComputerSystem.AutomaticManagedPagefile
		Wscript.Echo "AutomaticResetBootOption = " & objComputerSystem.AutomaticResetBootOption
		Wscript.Echo "AutomaticResetCapability = " & objComputerSystem.AutomaticResetCapability
		Wscript.Echo "BootOptionOnLimit = " & objComputerSystem.BootOptionOnLimit
		Wscript.Echo "BootOptionOnWatchDog = " & objComputerSystem.BootOptionOnWatchDog
		Wscript.Echo "BootROMSupported = " & objComputerSystem.BootROMSupported
		Wscript.Echo "BootupState = " & objComputerSystem.BootupState
		Wscript.Echo "Caption = " & objComputerSystem.Caption
		Wscript.Echo "ChassisBootupState = " & objComputerSystem.ChassisBootupState
		Wscript.Echo "CreationClassName = " & objComputerSystem.CreationClassName
		Wscript.Echo "CurrentTimeZone = " & objComputerSystem.CurrentTimeZone
		Wscript.Echo "DaylightInEffect = " & objComputerSystem.DaylightInEffect
		Wscript.Echo "Description = " & objComputerSystem.Description
		Wscript.Echo "DNSHostName = " & objComputerSystem.DNSHostName
		Wscript.Echo "Domain = " & objComputerSystem.Domain
		Wscript.Echo "DomainRole = " & objComputerSystem.DomainRole
		Wscript.Echo "EnableDaylightSavingsTime = " & objComputerSystem.EnableDaylightSavingsTime
		Wscript.Echo "FrontPanelResetStatus = " & objComputerSystem.FrontPanelResetStatus
		Wscript.Echo "InfraredSupported = " & objComputerSystem.InfraredSupported
		Wscript.Echo "InitialLoadInfo = " & objComputerSystem.InitialLoadInfo
		Wscript.Echo "InstallDate = " & objComputerSystem.InstallDate
		Wscript.Echo "KeyboardPasswordStatus = " & objComputerSystem.KeyboardPasswordStatus
		Wscript.Echo "LastLoadInfo = " & objComputerSystem.LastLoadInfo
		Wscript.Echo "Manufacturer = " & objComputerSystem.Manufacturer
		Wscript.Echo "Model = " & objComputerSystem.Model
		Wscript.Echo "Name = " & objComputerSystem.Name
		Wscript.Echo "NameFormat = " & objComputerSystem.NameFormat
		Wscript.Echo "NetworkServerModeEnabled = " & objComputerSystem.NetworkServerModeEnabled
		Wscript.Echo "NumberOfLogicalProcessors = " & objComputerSystem.NumberOfLogicalProcessors
		Wscript.Echo "NumberOfProcessors = " & objComputerSystem.NumberOfProcessors
		'Wscript.Echo "OEMLogoBitmap[] = " & objComputerSystem.OEMLogoBitmap[]
		'Wscript.Echo "OEMArray[] = " & objComputerSystem.OEMArray[]
		Wscript.Echo "PartOfDomain = " & objComputerSystem.PartOfDomain
		Wscript.Echo "PauseAfterReset = " & objComputerSystem.PauseAfterReset
		Wscript.Echo "PCSystemType = " & objComputerSystem.PCSystemType
		'Wscript.Echo "PowerManagementCapabilities[] = " & objComputerSystem.PowerManagementCapabilities[]
		Wscript.Echo "PowerManagementSupported = " & objComputerSystem.PowerManagementSupported
		Wscript.Echo "PowerOnPasswordStatus = " & objComputerSystem.PowerOnPasswordStatus
		Wscript.Echo "PowerState = " & objComputerSystem.PowerState
		Wscript.Echo "PowerSupplyState = " & objComputerSystem.PowerSupplyState
		Wscript.Echo "PrimaryOwnerContact = " & objComputerSystem.PrimaryOwnerContact
		Wscript.Echo "PrimaryOwnerName = " & objComputerSystem.PrimaryOwnerName
		Wscript.Echo "ResetCapability = " & objComputerSystem.ResetCapability
		Wscript.Echo "ResetCount = " & objComputerSystem.ResetCount
		Wscript.Echo "ResetLimit = " & objComputerSystem.ResetLimit
		'Wscript.Echo "Roles[] = " & objComputerSystem.Roles[]
		Wscript.Echo "Status = " & objComputerSystem.Status
		'Wscript.Echo "SupportContactDescription[] = " & objComputerSystem.SupportContactDescription[]
		Wscript.Echo "SystemStartupDelay = " & objComputerSystem.SystemStartupDelay
		'Wscript.Echo "SystemStartupOptions[] = " & objComputerSystem.SystemStartupOptions[]
		Wscript.Echo "SystemStartupSetting = " & objComputerSystem.SystemStartupSetting
		Wscript.Echo "SystemType = " & objComputerSystem.SystemType
		Wscript.Echo "ThermalState = " & objComputerSystem.ThermalState
		Wscript.Echo "TotalPhysicalMemory = " & objComputerSystem.TotalPhysicalMemory
		Wscript.Echo "UserName = " & objComputerSystem.UserName
		Wscript.Echo "WakeUpType = " & objComputerSystem.WakeUpType
		Wscript.Echo "Workgroup = " & objComputerSystem.Workgroup
	}
		
		
	   Next
End Sub

displaycomputersystem