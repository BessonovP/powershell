Import-Module ActiveDirectory #����� ����� ���� ������������ ������� ����� Get-ADUser � Set-ADUser

Import-Csv "addemail.csv" -Encoding UTF8 | ForEach-Object { #����������� CSV-����
    $displayName = $_.displayName #��������� �������� ������� displayName �� ������� ������ CSV
    $email = $_.EmailAddress #��������� �������� email-������ �� ������� ������ CSV

    # ����� �� DisplayName ����� ����������� � $displayName, ���� ������ ����������
    $users = @(Get-ADUser -Filter "DisplayName -eq '$displayName'" -Properties EmailAddress)

    if ($users.Count -eq 1) { #���� ������ ������ ���� ������������
        Set-ADUser -Identity $users[0].DistinguishedName -EmailAddress $email
        Write-Host "�������: $displayName => $email" -ForegroundColor Green #������������ �������
	Add-Content "update.txt" "�������: $displayName => $email" #��������� ����������� � �����
        }
    elseif ($users.Count -gt 1) { #���� ������ ����� ������ ������������
        Write-Host "��������� ������������� � ������: $displayName" -ForegroundColor Yellow #������������ ������
	$users | ForEach-Object { Write-Host " - $($_.DistinguishedName)" } #���������� ���������
	Add-Content "duplicate.txt" "���������: $users => $email" #��������� ��������� � �����
#	Add-Content "duplicate.txt" "���������: $($_.DistinguishedName) => $email" #��������� ��������� � �����

    }
    else { #���� �� ������� �� ������ ������������
        Write-Host "������������ �� ������: $displayName" -ForegroundColor Red #������������ �������
	Add-Content "fail.txt" "�����������: $displayName => $email" #��������� ��������� � �����
          }
}
