#Собираю все компьютеры домена в текстовый файл 
Get-ADComputer -SearchBase "OU=Employees, DC=acra-ratings, DC=ru" -Filter '*' | Select -Exp Name > .\computers.txt

#Загружаю файл со всеми компьютерами в переменную
$AllComputers = (Get-Content .\computers.txt)

#Объявление переменной для хранения всех ПК
$All_PC_info = @()

#Объявляение переменнй для хранения всех мониторов
$All_Monitor_info = @()

#Начинаю пробегаться циклом по каждому компьютеру 
ForEach($item in $AllComputers){
    
    #Проверяю пингуется ли ПК 
    if ( (Test-NetConnection $item -informationlevel quiet) )
    {
        
        #Включаю службу для удалённого выполнения PW 
        psexec \\$item sc start WinRm

        #Узнаю имя залогиненого пользователя
        [pscustomobject]$pc_user =  quser /server $item

        [string]$pc_user = ($pc_user[1].split(" "))[1]

        #Выгружаю из АД информацию о пользователе
        $norm_user_name = Get-ADUser $pc_user -properties extensionattribute11, extensionattribute12, extensionattribute13

        #Складываю ФИО пользователя на русском языке
        #[string]$norm_user_name = $norm_user_name.extensionattribute11 + " " + $norm_user_name.extensionattribute12 + " " + $norm_user_name.extensionattribute13
        
        #Узнаю информацию о ПК пользователя и выгружаю данные в файл  
        $bios_info = Get-CimInstance -ComputerName $item -ClassName win32_bios

        #Узанаю информацию о мониторах
        $monitors_info = Get-CimInstance -ComputerName $item -ClassName wmiMonitorID -Namespace root\wmi

        #Объявление переменной для хранения всех имён мониторов пользователя
        $all_name_mon = @()

        #Объявление переменной для хранения всех серийных номеров мониторов пользователя
        $all_serial_mon = @()
        
        #Цикл пробегается по списку мониторов (т.к. их может быть несколько)
        foreach ( $monitor in $monitors_info ){
            
            #Объявление переменных для хранения имени и серийного номера монитора
            $mon_name = ''
            $mon_serial = ''
            
            #Расшифровка имени и серийного номера монитора и запись их в переменные
            $monitor.UserFriendlyName | where { $_ -ne 0 } | foreach { $mon_name += [char]$_ }
            $monitor.SerialNumberID | where { $_ -ne 0 } | foreach { $mon_serial += [char]$_ }

            #Запись имени монитора и сериного номера в переменные
            $all_name_mon += $mon_name
            $all_serial_mon += $mon_serial

        }

        #Создание таблицы для хранения информации о пользователе и его мониторах
        $all_mon = @{

            "Пользователь" = [string]$pc_user;  
            "Наименования" = [string]$all_name_mon;
            "Серийный номер" = [string]$all_serial_mon

        }

        #Запись информации о мониторах конкретного пользователя в переменную со всеми данными
        $All_Monitor_info += [pscustomobject]$all_mon

        #Запись информации о ПК пользователя в переменную со всеми данными ( + преобразование в читаемый вид — таблицу )
        $All_PC_info += $bios_info | select @{ "Name" = "Пользователь"; "expression" = { $pc_user } }, 
            @{ "Name" = "ФИО (если есть)"; "expression" = { [string]$norm_user_name } },
            @{ "Name" = "Имя ПК"; "expression" = { [string]$item } },
            @{ "Name" = "Производитель"; "expression" = { $_.Manufacturer } }, 
            @{ "Name" = "Серийный номер ПК"; "expression" = { $_.SerialNumber } }


        #Отключаю службу для удалённого выполнения PW 
        psexec \\$item sc stop WinRm

    }
    else
    {
        
        #Если ПК не пингуется добавляю его в отдельный файл 
        $item | Out-File -FilePath .\NoPingPC.txt -Append

    }

}

#Экспорт данных о всех ПК в CSV с учётом языка и кодировки
$All_PC_info | Export-Csv .\inventory_PC.csv -notype -UseCulture -Encoding UTF8

#Экспорт данных о всех мониторах в CSV с учётом языка и кодировки
$All_Monitor_info | Export-Csv .\inventory_Monitors.csv -notype -UseCulture -Encoding UTF8