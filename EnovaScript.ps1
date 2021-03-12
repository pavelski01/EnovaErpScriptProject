# Usage: .\EnovaScript.ps1

# Add Soneta types
Add-Type -Path .\Soneta.Business.dll
Add-Type -Path .\Soneta.Start.dll
Add-Type -Path .\Soneta.Types.dll

try {
    # Loader
    $loader = [Soneta.Start.Loader]::new()
    $loader.WithUI = $true
    $loader.WithNet = $false
    $loader.WithExtra = $true
    $loader.WithExtensions = $false
    $loader.Load()

    # Database
    $database = [Soneta.Business.App.MsSqlDatabase]::new()
    $database.Name = 'Nowa_firma'
    $database.Server = '.\enova'
    $database.DatabaseName = 'Nowa_firma'
    $database.User = 'sa'
    $database.Password = 'qwerty12345'
    $database.Trusted = $false
    $database.Active = $true

    # ERP account credentials
    $loginParameters = [Soneta.Business.App.LoginParameters]::new()
    $loginParameters.Operator = 'Administrator'
    $loginParameters.Password = ''
    $loginParameters.Mode = [Soneta.Types.AuthenticationType]::UserPassword

    # Connect to database
    $login = $database.Login($loginParameters)

    # Initialize program
    $customAttributes = `
        [Soneta.Tools.AssemblyAttributes]::GetCustom([Soneta.Business.ProgramInitializerAttribute])
    foreach ($ca in $customAttributes) {
        $pia = [Soneta.Business.ProgramInitializerAttribute]$ca
        $pi = `
            [Soneta.Business.IProgramInitializer]$pia.InitializerType.GetConstructor(`
                [System.Type]::EmptyTypes`
            ).Invoke($null)
        $pi.Initialize()
    }

    # Begin session
    $session = `
        $login.CreateSession($false, $false, 'ErpSession: ' + [System.Guid]::NewGuid().ToString('N').ToUpper())  
    
    # Initialize modules
    $tm = [Soneta.Handel.HandelModule]::GetInstance($session)
    $crm = [Soneta.CRM.CRMModule]::GetInstance($session)
    $wm = [Soneta.Magazyny.MagazynyModule]::GetInstance($session)
    $mm = [Soneta.Towary.TowaryModule]::GetInstance($session)

    # Begin upper transaction
    $documentTransaction = $session.Logout($true)
    $document = [Soneta.Handel.DokumentHandlowy]::new()
    $document.Definicja = $tm.DefDokHandlowych.WgSymbolu['FV']
    $document.Magazyn = $wm.Magazyny.Firma;
    $tm.DokHandlowe.AddRow($document);
    $document.Kontrahent = $crm.Kontrahenci.WgKodu['ABC']
    $merchandise = $mm.Towary.WgKodu['TRANSPORT']
    
    # Begin under transation
    $positionTransaction = $session.Logout($true)
    $position = [Soneta.Handel.PozycjaDokHandlowego]::new($document)
    $tm.PozycjeDokHan.AddRow($position)
    $position.Towar = $merchandise
    $position.Ilosc = [Soneta.Towary.Quantity]::new(10, $null)
    $position.Cena = [Soneta.Types.DoubleCy]::new(100.0)
    $positionTransaction.CommitUI()
    $positionTransaction.Dispose()
    # End under transation

    $document.Stan = [Soneta.Handel.StanDokumentuHandlowego]::Zatwierdzony
    $documentTransaction.Commit()
    $documentTransaction.Dispose()
    # End upper transation

    $session.Save()
    $session.Dispose()
    # End session
}
catch {    
    Write-Host -Foreground Red $_.Exception
}