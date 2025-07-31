Import-Module ActiveDirectory #чтобы можно было использовать команды вроде Get-ADUser и Set-ADUser

Import-Csv "addemail.csv" -Encoding UTF8 | ForEach-Object { #Импортирует CSV-файл
    $displayName = $_.displayName #Извлекает значение столбца displayName из текущей строки CSV
    $email = $_.EmailAddress #Извлекает значение email-адреса из текущей строки CSV

    # Поиск по DisplayName точно совпадающее с $displayName, ищет точное совпадение
    $users = @(Get-ADUser -Filter "DisplayName -eq '$displayName'" -Properties EmailAddress)

    if ($users.Count -eq 1) { #если найден только один пользователь
        Set-ADUser -Identity $users[0].DistinguishedName -EmailAddress $email
        Write-Host "Обновлён: $displayName => $email" -ForegroundColor Green #подсвечивает зеленым
	Add-Content "update.txt" "Обновлён: $displayName => $email" #добавляет обновленных в отчет
        }
    elseif ($users.Count -gt 1) { #если найден более одного пользователя
        Write-Host "Несколько пользователей с именем: $displayName" -ForegroundColor Yellow #подсвечивает желтым
	$users | ForEach-Object { Write-Host " - $($_.DistinguishedName)" } #показывает дубликаты
	Add-Content "duplicate.txt" "Дубликаты: $users => $email" #добавляет дубликаты в отчет
#	Add-Content "duplicate.txt" "Дубликаты: $($_.DistinguishedName) => $email" #добавляет дубликаты в отчет

    }
    else { #если не найдено ни одного пользователя
        Write-Host "Пользователь не найден: $displayName" -ForegroundColor Red #подсвечивает красным
	Add-Content "fail.txt" "Отсутствует: $displayName => $email" #добавляет дубликаты в отчет
          }
}
