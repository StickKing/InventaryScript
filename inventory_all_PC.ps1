#Выставляем OU-шку со всеми ПК компании и выгружаем их имена
$AllComputers = Get-ADComputer -SearchBase "OU=Employees, DC=YouDomain, DC=com" -Filter '*' | Select -Exp Name

#Определяем пустой список для хранения информации о ПК
$All_PC_info = @()

#Определяем пустой список для хранения информации о мониторах
$All_Monitor_info = @()

#Проходим по всем выгруженым из AD ПК
ForEach($item in $AllComputers){
    
    #Проверяем находится ли ПК в сети (пингуем по факту)
    if ( (Test-NetConnection $item -informationlevel quiet) )
    {
        
        #НЕОБЯЗАТЕЛЬНО: Если служба WinRm выключена, то включаем её сторонней утилитой psexec
        #psexec \\$item sc start WinRm

        #Получаем информацию о залогиненых пользователях ПК
        [pscustomobject]$pc_user =  quser /server $item
        #Вырезаем имя пользователя
        [string]$pc_user = ($pc_user[1].split(" "))[1]

        #Выгружаем из AD информацию о залогиненом пользователе (выгружаем атрибуты extensionattribute11, extensionattribute12, extensionattribute13 если там записана полезная информация)
        $norm_user_name = Get-ADUser $pc_user -properties extensionattribute11, extensionattribute12, extensionattribute13

        #НЕОБЯЗАТЕЛЬНО: Если в extensionattribute11, extensionattribute12, extensionattribute13 записано ФИО сотруника то преобразуем его в строку
        #[string]$norm_user_name = $norm_user_name.extensionattribute11 + " " + $norm_user_name.extensionattribute12 + " " + $norm_user_name.extensionattribute13
        
        #Через WMI объекты получаем информацию о ПК 
        $bios_info = Get-CimInstance -ComputerName $item -ClassName win32_bios

        #Через WMI объекты получаем информацию о мониторах
        $monitors_info = Get-CimInstance -ComputerName $item -ClassName wmiMonitorID -Namespace root\wmi

        #Создаём список для имён мониторов т.к. их может быть несколько у одного ПК
        $all_name_mon = @()
        #Создаём список для серийных номеров мониторов т.к. их может быть несколько у одного ПК
        $all_serial_mon = @()
        
        #Проходимся по всем монитором присоединёным к ПК
        foreach ( $monitor in $monitors_info ){
            
            #Локальные переменные для имени и серийника
            $mon_name = ''
            $mon_serial = ''
            
            #Получаем имя и серийный номер монитора
            $monitor.UserFriendlyName | where { $_ -ne 0 } | foreach { $mon_name += [char]$_ }
            $monitor.SerialNumberID | where { $_ -ne 0 } | foreach { $mon_serial += [char]$_ }

            #Добавляем полученые данные в списки имён и серийных номеров
            $all_name_mon += $mon_name
            $all_serial_mon += $mon_serial

        }

        #Преобразуем полченные данные в табличный вид
        $all_mon = @{
            "Пользователь" = [string]$pc_user;  
            "Наименование" = [string]$all_name_mon;
            "Серийный номер" = [string]$all_serial_mon
        }

        #Добавляем мониторы в общий список мониторов
        $All_Monitor_info += [pscustomobject]$all_mon

        #Добавляем информацию о ПК в общий список ПК
        $All_PC_info += $bios_info | select @{ "Name" = "Пользователь"; "expression" = { $pc_user } }, 
            @{ "Name" = "ФИО (если есть)"; "expression" = { [string]$norm_user_name } },
            @{ "Name" = "Имя ПК"; "expression" = { [string]$item } },
            @{ "Name" = "Производитель"; "expression" = { $_.Manufacturer } }, 
            @{ "Name" = "Серийный номер ПК"; "expression" = { $_.SerialNumber } }


        #НЕОБЯЗАТЕЛЬНО: Если необходимо то выключаем службу WinRm с помощью psexec
        #psexec \\$item sc stop WinRm

    }
    else
    {
        #ПК вне сети (которые не пинговались) записываем в текстовый файл
        $item | Out-File -FilePath .\NoPingPC.txt -Append
    }

}

#Преобразуем информацию о ПК в файл CSV с поддержкой русского языка
$All_PC_info | Export-Csv .\inventory_PC.csv -notype -UseCulture -Encoding UTF8

##Преобразуем информацию о мониторах в файл CSV с поддержкой русского языка
$All_Monitor_info | Export-Csv .\inventory_Monitors.csv -notype -UseCulture -Encoding UTF8