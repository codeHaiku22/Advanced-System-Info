   M S B A M L   `   `   `  ���� E  AAdvSysInfo, Version=1.0.2.0, Culture=neutral, PublicKeyToken=null'clr-namespace:AdvSysInfo
AdvSysInfo      AdvSysInfo.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     *localclr-namespace:AdvSysInfo       � Title$  
AdvSysInfo��$
��1040��$
��1920�� � WindowState$ 	Maximized=�.��(� .��� 
�������.+�� -    �Name$ 
grdSysInfo��$	��300��$	��Top=�.+�� -   $ grdSpecs��$	��300��$	��500��$	��Top=�$
��Left=�.+��� $��Host Name/IP:��$
��Left=�$��28��$��	10,10,0,0q�$	��Top=��� -    ��Name$ txtServerNameIP��$
��Left=�$��28��$��
108,10,0,0q� �� TextWrapping$
 Wrap=�$	��Top=�$��Center=�$	��302���� -    ��Name$ lblServerNameIP��$��lblServerNameIP��$
��Left=�$��
108,34,0,0q�$	��Top=��� -    ��Name$ cmdScan��$
��Scan��$
��Left=�$��
415,12,0,0q�$	��Top=�$��75��$��23���� $��System Specs��$
��Left=�$	��225��$��	10,65,0,0q�$	��Top=�$	��480��
�������.���� -    ��Name$ 
lbxDetails��$
��Left=�$	��200��$��0,0,0,0q�$��Bottom=�$	��465��
�������$��12(��� $��	Processor��$
��Left=�$	��210��$��	500,5,0,0q�$	��Top=�$	��320��$��5q�
�������.���� -    ��Name$ scrlCPU��.��� -   $ grdProc��$	��200��$	��305��$��Bottom=�$
��Left=�$��0,0,0,0q�	 � ColumnDefinitions	 �� 
 � RowDefinitions
 �� �� $��Memory��$
��Left=�$��65��$��500,225,0,0q�$	��Top=�$	��320��$��5q�
�������.��� -	   $ grdMem��$��45��$	��305��$��Center=�$
��Left=�$��0,0,0,0q�	 �� 
 �� �� $��Disks��$
��Left=�$	��285��$��	830,5,0,0q�$	��Top=�$
��1065��$��0q�
�������.���� -
   $ 	scrlDisks��.��� -   $ grdDisks��$	��270��$
��1060��$��Bottom=�$
��Left=�$��0,0,0,0q�	 �� �� �� �� 
 �� �� �� �� �� -    ��Name$ tabctrl��$
��Left=�$	��680��$��
10,299,0,0q�$	��Top=�$
��1885��.@��� -    ��Name$ tabApplications��$��Applications��.��� ���	#FFE5E5E5.+��� -    ��Name$ stkpnlApplications�� �� Orientation$ Vertical=�.+��� -   $  stkpnlApplicationsCriteria��$ 
Horizontal=�.+��� -    ��Name$ grpApplicationName��$��Application��$
��Left=�$��65��$��	10,10,0,0q�$	��Top=�$	��340��.���� -   $ stkpnlApplicationName��$ 
Horizontal=�.+��� $��Name:��$
��Left=�$��10,5,0,0q�$	��Top=��� -   $ txtApplicationName��$
��Left=�$��	10,10,0,0q�$	��Top=�$	��250���� -   $ grpBit��$��Environment��$
��Left=�$��65��$��	25,10,0,0q�$	��Top=�$	��165��.���� -   $ 	stkpnlBit��$ 
Horizontal=�.+��� -    ��Name$ chk32Bit��$��32-Bit��$
��Left=�$��	10,10,0,0q�$	��Top=��� -   $ chk64Bit��$��64-Bit��$
��Left=�$��	15,10,0,0q�$	��Top=��� -   $ cmdSearchApplications��$��Search��$
��Left=�$��	25,35,0,0q�$	��Top=�$��75��$��23���� -   $ lblApplicationsCount��$�� ��$
��Left=�$��	25,35,0,0q�$	��Top=�&   System.Windows.Controls.DataGrid  -     Name$ dtgrdApplications��$
��Left=�$	��547��$��	10,10,0,0q�$	��Top=�$
��1859������ �� -   $ tabProcesses��$��	Processes��.��� ���	#FFE5E5E5.+��� -   $ stkpnlProcesses��$ Vertical=�.+��� -   $ stkpnlProcessesCriteria��$ 
Horizontal=�.+��� -   $ grpProcessName��$��Processes Filter��$
��Left=�$��65��$��	10,10,0,0q�$	��Top=�$	��350��.���� -   $ stkpnlProcessName��$ 
Horizontal=�.+��� $��Name:��$
��Left=�$��10,5,0,0q�$	��Top=��� -   $ txtProcessName��$
��Left=�$��	10,10,0,0q�$	��Top=�$	��250���� -    $ cmdSearchProcesses��$��Search��$
��Left=�$��	25,35,0,0q�$	��Top=�$��75��$��23���� -!   $ lblProcessesCount��$�� ��$
��Left=�$��	25,35,0,0q�$	��Top=�  -"   $ dtgrdProcesses��$
��Left=�$	��547��$��	10,10,0,0q�$	��Top=�$
��1859������ �� -#   $ tabServices��$��Services��.��� ���	#FFE5E5E5.+��� -$   $ stkpnlServices��$ Vertical=�.+��� -%   $ stkpnlServicesCriteria��$ 
Horizontal=�.+��� -&   $ grpServiceType��$��Service Type��$
��Left=�$��65��$��	10,10,0,0q�$	��Top=�$	��220��.���� -'   $ stkpnlServiceType��$ 
Horizontal=�.+�	� -(    	�Name$ 	rbGeneral��$��General��$
��Left=�$��	10,10,0,0q�$	��Top=� 	� 	GroupName$ serviceType��	� -)   $ rbDeviceDriver��$��Device Driver��$
��Left=�$��	15,10,0,0q�$	��Top=�$ serviceType���� -*   $ grpServiceStatus��$��Status��$
��Left=�$��65��$��	25,10,0,0q�$	��Top=�$	��200��.���� -+   $ stkpnlServiceStatus��$ 
Horizontal=�.+��� -,   $ 
chkRunning��$��Running��$
��Left=�$��	10,10,0,0q�$	��Top=��� --   $ 
chkStopped��$��Stopped��$
��Left=�$��	15,10,0,0q�$	��Top=��� -.   $ 
grpStartup��$��Startup��$
��Left=�$��65��$��	25,10,0,0q�$	��Top=�$	��305��.���� -/   $ stkpnlStartup��$ 
Horizontal=�.+��� -0   $ chkAutomatic��$��	Automatic��$
��Left=�$��	10,10,0,0q�$	��Top=��� -1   $ 	chkManual��$��Manual��$
��Left=�$��	20,10,0,0q�$	��Top=��� -2   $ chkDisabled��$��Disabled��$
��Left=�$��	20,10,0,0q�$	��Top=��� -3   $ cmdSearchServices��$��Search��$
��Left=�$��	25,35,0,0q�$	��Top=�$��75��$��23���� -4   $ lblServicesCount��$�� ��$
��Left=�$��	25,35,0,0q�$	��Top=�  -5   $ dtgrdServices��$
��Left=�$	��547��$��	10,10,0,0q�$	��Top=�$
��1859������ �� -6   $ tabEventLog��$��
Event Logs��.��� ���	#FFE5E5E5.+��� -7   $ stkpnlEventLogs��$ Vertical=�.+��� -8   $ stkpnlEventLogCriteria��$ 
Horizontal=�.+��� -9   $ grpEventLogType��$��Event Log Type��$
��Left=�$��65��$��	10,10,0,0q�$	��Top=�$	��300��.���� -:   $ stkpnlEventLogType��$ 
Horizontal=�.+�	� -;   $ rbApplication��$��Application Events��$
��Left=�$��	10,10,0,0q�$	��Top=�$ logType��	� -<   $ rbSystem��$��System Events��$
��Left=�$��	15,10,0,0q�$	��Top=�$ logType���� -=   $ grpEventLevel��$��Event Level��$
��Left=�$��65��$��	25,10,0,0q�$	��Top=�$	��470��.���� ->   $ stkpnlEventLevel��$ 
Horizontal=�.+��� -?   $ chkCritical��$��Critical��$
��Left=�$��	10,10,0,0q�$	��Top=��� -@   $ 
chkWarning��$��Warning��$
��Left=�$��	20,10,0,0q�$	��Top=��� -A   $ 
chkVerbose��$��Verbose��$
��Left=�$��	20,10,0,0q�$	��Top=��� -B   $ chkError��$��Error��$
��Left=�$��	20,10,0,0q�$	��Top=��� -C   $ chkInformation��$��Information��$
��Left=�$��	20,10,0,0q�$	��Top=��� -D   $ grpDateRange��$��
Date Range��$
��Left=�$��65��$��	25,10,0,0q�$	��Top=�$	��425��.���� -E   $ stkpnlDateRange��$ 
Horizontal=�.+��� $��
From Date:��$
��Left=�$��10,5,0,0q�$	��Top=�(  "System.Windows.Controls.DatePicker  -F     Name$ dtpFromDate��$
��Left=�$��0,5,0,0q�$	��Top=��� $��To Date:��$
��Left=�$��20,5,0,0q�$	��Top=�  -G   $ 	dtpToDate��$
��Left=�$��0,5,0,0q�$	��Top=��� -H   $ cmdSearchEventLogs��$��Search��$
��Left=�$��	25,35,0,0q�$	��Top=�$��75��$��23���� -I   $ lblEventLogCount��$�� ��$
��Left=�$��	25,35,0,0q�$	��Top=�  -J   $ dtgrdEventLogs��$
��Left=�$	��390��$��	10,10,0,0q�$	��Top=�$
��1859������ �� -K   $ grpEventDetails��$��Event Details��$
��Left=�$	��155��$��	10,10,0,0q�$	��Top=�$
��1859��.���� -L   $ stkpnlEventDetails��$ Vertical=�.+��� -M   $ stkpnlEventDetails1��$ 
Horizontal=�.+��� -N   $ lblEventIdCaption��$��	Event Id:��$
��Left=�$��10,5,0,0q�$	��Top=��� -O   $ 
lblEventId��$
��Left=�$��0,5,0,0q�$	��Top=��� -P   $ lblEventTypeCaption��$��Event Type:��$
��Left=�$��50,5,0,0q�$	��Top=��� -Q   $ lblEventType��$
��Left=�$��0,5,0,0q�$	��Top=��� -R   $ lblGeneratedCaption��$��
Generated:��$
��Left=�$��50,5,0,0q�$	��Top=��� -S   $ lblGenerated��$
��Left=�$��0,5,0,0q�$	��Top=��� -T   $ lblSourceCaption��$��Event Source:��$
��Left=�$��50,5,0,0q�$	��Top=��� -U   $ 	lblSource��$
��Left=�$��0,5,0,0q�$	��Top=��� -V   $ lblMachineNameCaption��$��Machine Name:��$
��Left=�$��50,5,0,0q�$	��Top=��� -W   $ lblMachineName��$
��Left=�$��0,5,0,0q�$	��Top=��� -X   $ stkpnlEventDetails2��$ 
Horizontal=�.+��� -Y   $ lblMessageCaption��$��Message:��$
��Left=�$��10,0,0,0q�$	��Top=��� -Z   $ 
txtMessage��$ WrapWithOverflow=�$
��Left=�$	��Top=�$��0,5,0,0q�$
��1775��$��90��$��Visible=� �� 
IsReadOnly . ���	#FFE5E5E5�� -[   $ tabOpenFiles��$��
Open Files��.��� ���	#FFE5E5E5.+��� -\   $ stkpnlOpenFiles��$ Vertical=�.+��� -]   $ stkpnlOpenFilesCriteria��$ 
Horizontal=�.+��� -^   $ 	grpFilter��$��Open Files Filter��$
��Left=�$	��100��$��	10,10,0,0q�$	��Top=�$	��510��.���� -_   $ stkpnlFilter��$ 
Horizontal=�.+��� -`   $ stkpnlFields��$ Vertical=�.+��� -a   $ stkpnlAccessedBy��$ 
Horizontal=�.+��� -b   $ chkAccessedBy��$��Accessed By:��$
��Left=�$��	10,10,0,0q�$	��Top=��� -c   $ txtAccessedBy��$
��Left=�$��	10,10,0,0q�$	��Top=�$	��250���� -d   $ stkpnlOpenFile��$ 
Horizontal=�.+��� -e   $ chkOpenFile��$��Open File:    ��$
��Left=�$��	10,10,0,0q�$	��Top=��� -f   $ txtOpenFile��$
��Left=�$��	10,10,0,0q�$	��Top=�$	��250���� -g   $ stkpnlAndOr��$ 
Horizontal=�.+�	� -h   $ rbAnd��$	��And��$��Center=�$��20,0,0,0q�$��Center=�$ andOr��	� -i   $
 rbOr��$��Or��$��Center=�$��10,0,0,0q�$��Center=�$ andOr���� -j   $ cmdSearchOpenFiles��$��Search��$
��Left=�$��	25,20,0,0q�$��Center=�$��75��$��23���� -k   $ lblOpenFilesCount��$�� ��$
��Left=�$��	25,20,0,0q�$��Center=�  -l   $ dtgrdOpenFiles��$
��Left=�$	��515��$��	10,10,0,0q�$	��Top=�$
��1859������ 