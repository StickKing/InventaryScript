
$AllComputers = Get-ADComputer -SearchBase "OU=Employees, DC=YouDomain, DC=com" -Filter '*' | Select -Exp Name

#���������� ���������� ��� �������� ���� ��
$All_PC_info = @()

#����������� ��������� ��� �������� ���� ���������
$All_Monitor_info = @()

#������� ����������� ������ �� ������� ���������� 
ForEach($item in $AllComputers){
    
    #�������� ��������� �� �� 
    if ( (Test-NetConnection $item -informationlevel quiet) )
    {
        
        #������� ������ ��� ��������� ���������� PW 
        psexec \\$item sc start WinRm

        #����� ��� ������������ ������������
        [pscustomobject]$pc_user =  quser /server $item

        [string]$pc_user = ($pc_user[1].split(" "))[1]

        #�������� �� �� ���������� � ������������
        $norm_user_name = Get-ADUser $pc_user -properties extensionattribute11, extensionattribute12, extensionattribute13

        #��������� ��� ������������ �� ������� �����
        #[string]$norm_user_name = $norm_user_name.extensionattribute11 + " " + $norm_user_name.extensionattribute12 + " " + $norm_user_name.extensionattribute13
        
        #����� ���������� � �� ������������ � �������� ������ � ����  
        $bios_info = Get-CimInstance -ComputerName $item -ClassName win32_bios

        #������ ���������� � ���������
        $monitors_info = Get-CimInstance -ComputerName $item -ClassName wmiMonitorID -Namespace root\wmi

        #���������� ���������� ��� �������� ���� ��� ��������� ������������
        $all_name_mon = @()

        #���������� ���������� ��� �������� ���� �������� ������� ��������� ������������
        $all_serial_mon = @()
        
        #���� ����������� �� ������ ��������� (�.�. �� ����� ���� ���������)
        foreach ( $monitor in $monitors_info ){
            
            #���������� ���������� ��� �������� ����� � ��������� ������ ��������
            $mon_name = ''
            $mon_serial = ''
            
            #����������� ����� � ��������� ������ �������� � ������ �� � ����������
            $monitor.UserFriendlyName | where { $_ -ne 0 } | foreach { $mon_name += [char]$_ }
            $monitor.SerialNumberID | where { $_ -ne 0 } | foreach { $mon_serial += [char]$_ }

            #������ ����� �������� � �������� ������ � ����������
            $all_name_mon += $mon_name
            $all_serial_mon += $mon_serial

        }

        #�������� ������� ��� �������� ���������� � ������������ � ��� ���������
        $all_mon = @{

            "������������" = [string]$pc_user;  
            "������������" = [string]$all_name_mon;
            "�������� �����" = [string]$all_serial_mon

        }

        #������ ���������� � ��������� ����������� ������������ � ���������� �� ����� �������
        $All_Monitor_info += [pscustomobject]$all_mon

        #������ ���������� � �� ������������ � ���������� �� ����� ������� ( + �������������� � �������� ��� � ������� )
        $All_PC_info += $bios_info | select @{ "Name" = "������������"; "expression" = { $pc_user } }, 
            @{ "Name" = "��� (���� ����)"; "expression" = { [string]$norm_user_name } },
            @{ "Name" = "��� ��"; "expression" = { [string]$item } },
            @{ "Name" = "�������������"; "expression" = { $_.Manufacturer } }, 
            @{ "Name" = "�������� ����� ��"; "expression" = { $_.SerialNumber } }


        #�������� ������ ��� ��������� ���������� PW 
        psexec \\$item sc stop WinRm

    }
    else
    {
        
        #���� �� �� ��������� �������� ��� � ��������� ���� 
        $item | Out-File -FilePath .\NoPingPC.txt -Append

    }

}

#������� ������ � ���� �� � CSV � ������ ����� � ���������
$All_PC_info | Export-Csv .\inventory_PC.csv -notype -UseCulture -Encoding UTF8

#������� ������ � ���� ��������� � CSV � ������ ����� � ���������
$All_Monitor_info | Export-Csv .\inventory_Monitors.csv -notype -UseCulture -Encoding UTF8