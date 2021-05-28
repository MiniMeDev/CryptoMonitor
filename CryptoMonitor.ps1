Clear-Host

#Function scrup data from CoinMarket website based on filters listed in $Parameters
function GetData {
    Param (
		[Parameter(Mandatory = $true)][string]$URL,
        [Parameter(Mandatory = $true)][string][string]$APIKey
	)
	process
	{ 
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("X-CMC_PRO_API_KEY", $APIKey)
    $headers.Add('Accepts', 'application/json')

    $Parameters = @{
        start = 1
        limit = 100
        convert = 'USD'
        price_max = 3
        sort = 'percent_change_1h'
        sort_dir = 'desc'
        percent_change_24h_min = 50
        volume_24h_min = 1000000
        volume_24h_max	= 100000000
      }

    $Rsult = Invoke-RestMethod -Uri $URL -Method Get -Headers $headers -Body $Parameters
    return $Rsult
    }
}


#Function generating report
function GetReport {
    param (
        [Parameter(Mandatory = $true)][PsObject]$Data
    )
    process
	{ 
        $Report=@()
        foreach ($Coin in $Data) {
            if ($Coin -and ($coin.Symbol -notlike '*DOWN') -and ($coin.Symbol -notlike '*BEAR')){

                $TwitterLink=$null; $TwitterLink = ($(Invoke-WebRequest "https://coinmarketcap.com/currencies/$($Coin.slug)/").Links  | Where-Object {$_.Class -eq "modalLink___MQefI"}  | Where-Object {$_.outerText -eq "Twitter"}).href
                #if([int](([System.Net.WebRequest]::Create($TwitterLink)).GetResponse().StatusCode) -ne 200 ){
               #   $TwitterLinkURL='Twitter'
               # }else{
                  #$TwitterLinkURL= "<a href=$($TwitterLink)/>Twitter</a>"
                  $TwitterLinkURL= "$($TwitterLink)"
                #}
                
                $CoinMarketLink = "https://coinmarketcap.com/currencies/$($Coin.slug)/"
                #if([int](([System.Net.WebRequest]::Create($CoinMarketLink)).GetResponse().StatusCode) -ne 200 ){
                 # $CoinMarketLinkURL='CoinMarket'
                #}else{
                 # $CoinMarketLinkURL= "<a href=$($CoinMarketLink)/>CoinMarket</a>"
                  $CoinMarketLinkURL= "$($CoinMarketLink)"
                #}
 
                if($Coin.platform.Name -like '*Ethereum*'){
                   # if([int](([System.Net.WebRequest]::Create("https://dex.guru/token/$($Coin.platform.token_address)-eth")).GetResponse().StatusCode) -ne 200 )
                   # {
                    #    $GuruURL='Guru'
                    #}else{
                       # $GuruURL= "<a href=https://dex.guru/token/$($Coin.platform.token_address)-eth/>Guru</a>"
                        $GuruURL= "https://dex.guru/token/$($Coin.platform.token_address)-eth"
                    #}
                }elseif($Coin.platform.name -like '*Binance*'){
                    #if([int](([System.Net.WebRequest]::Create("https://dex.guru/token/$($Coin.platform.token_address)-bsc")).GetResponse().StatusCode) -ne 200 )
                    #{
                        $GuruURL='Guru'
                    #}else{
                       # $GuruURL= "<a href=https://dex.guru/token/$($Coin.platform.token_address)-bsc/>Guru</a>"
                        $GuruURL= "https://dex.guru/token/$($Coin.platform.token_address)-bsc"
                    #}
                }
                $Info = New-Object -TypeName psobject
                $Info | Add-Member -MemberType NoteProperty -Name 'Added on' -Value ( get-date $Coin.date_added -Format dd-MM-yyyy)
                $Info | Add-Member -MemberType NoteProperty -Name 'Coin' -Value $Coin.Symbol
                $Info | Add-Member -MemberType NoteProperty -Name 'Name' -Value $Coin.name
                $Info | Add-Member -MemberType NoteProperty -Name 'Price' -Value ([math]::Round($Coin.quote.USD.Price,2))
                $Info | Add-Member -MemberType NoteProperty -Name 'Vol 24hr' -Value "$([math]::Round($Coin.quote.USD.volume_24h/1000000,2)) Mil"
                $Info | Add-Member -MemberType NoteProperty -Name 'Change % 1hr' -Value ([math]::Round($Coin.quote.USD.percent_change_1h,0))
                $Info | Add-Member -MemberType NoteProperty -Name 'Change % 24hr' -Value ([math]::Round($Coin.quote.USD.percent_change_24h,0))
                $Info | Add-Member -MemberType NoteProperty -Name 'Change % 7D' -Value ([math]::Round($Coin.quote.USD.percent_change_7d,0))
                $Info | Add-Member -MemberType NoteProperty -Name 'Market Cap' -Value "$([math]::Round($Coin.quote.USD.market_cap/1000000,2)) Mil"
                $Info | Add-Member -MemberType NoteProperty -Name 'Max Supply' -Value "$([math]::Round($Coin.max_supply/1000000,2)) Mil"
                $Info | Add-Member -MemberType NoteProperty -Name 'Total Supply' -Value "$([math]::Round($Coin.total_supply/1000000,2)) Mil"
                $Info | Add-Member -MemberType NoteProperty -Name 'Crculating Supply' -Value "$([math]::Round($Coin.circulating_supply/1000000,2)) Mil"
                $Info | Add-Member -MemberType NoteProperty -Name 'Token Address' -Value $Coin.platform.token_address
                $Info | Add-Member -MemberType NoteProperty -Name 'Block Chain' -Value $Coin.platform.name
                $Info | Add-Member -MemberType NoteProperty -Name 'Links' -Value ("$($CoinMarketLinkURL), $($TwitterLinkURL), $($GuruURL)")
                $Report +=$Info
            }
        }
    Return $Report
    }   
}

Function CredentialObj{
	Param (
		[Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$UserName,
		[Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][String]$Password,
		[String]$LogFilePath
		)
	Process {
		$Secpwd = ConvertTo-SecureString $Password -AsPlainText -Force
		$Cred = New-Object System.Management.Automation.PSCredential ($UserName, $secpwd)
		Return , $Cred
	}
}

#Css report table style
Function CssStyle{
	param(
		[Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][ValidateSet('Table','FlorAntara')][string]$Style
		)
	Process {
		switch ($Style)
		{
			Table { 
					$CSS ="@import url(http://fonts.googleapis.com/css?family=Roboto:400,500,700,300,100);

					body {
					  #background-color: #65abf2;
					  font-family: 'Roboto', helvetica, arial, sans-serif;
					  font-size: 16px;
					  font-weight: 400;
					  text-rendering: optimizeLegibility;
					}

					div.table-title {
					   display: block;
					  margin: auto;
					  max-width: 600px;
					  padding:5px;
					  width: 100%;
					}

					.table-title h3 {
					   color: #fafafa;
					   font-size: 30px;
					   font-weight: 400;
					   font-style:normal;
					   font-family: 'Roboto', helvetica, arial, sans-serif;
					   text-shadow: -1px -1px 1px rgba(0, 0, 0, 0.1);
					   text-transform:uppercase;
					}


					/*** Table Styles **/

					.table-fill {
					  background: white;
					  border-radius:3px;
					  border-collapse: collapse;
					  height: 320px;
					  margin: auto;
					  max-width: 600px;
					  padding:5px;
					  width: 100%;
					  box-shadow: 0 5px 10px rgba(0, 0, 0, 0.1);
					  animation: float 5s infinite;
					}
 
					th {
					  color:#D5DDE5;;
					  background:#1b1e24;
					  border-bottom:4px solid #9ea7af;
					  border-right: 1px solid #343a45;
					  font-size:14px;
					  font-weight: 100;
					  padding:12px;
					  text-align:left;
					  text-shadow: 0 1px 1px rgba(0, 0, 0, 0.1);
					  vertical-align:middle;
					}

					th:first-child {
					  border-top-left-radius:3px;
					}
 
					th:last-child {
					  border-top-right-radius:3px;
					  border-right:none;
					}
  
					tr {
					  border-top: 1px solid #C1C3D1;
					  border-bottom-: 1px solid #C1C3D1;
					  color:#666B85;
					  font-size:16px;
					  font-weight:normal;
					  text-shadow: 0 1px 1px rgba(256, 256, 256, 0.1);
					}
 
					tr:hover td {
					  background:#8dd2d1;
					  color:#FFFFFF;
					  border-top: 1px solid #22262e;
					  border-bottom: 1px solid #22262e;
					}
 
					tr:first-child {
					  border-top:none;
					}

					tr:last-child {
					  border-bottom:none;
					}
 
					tr:nth-child(odd) td {
					  background:#EBEBEB;
					}
 
					tr:nth-child(odd):hover td {
					  background:#8dd2d1;
					}

					tr:last-child td:first-child {
					  border-bottom-left-radius:3px;
					}
 
					tr:last-child td:last-child {
					  border-bottom-right-radius:3px;
					}
 
					td {
					  background:#FFFFFF;
					  padding:10px;
					  text-align:left;
					  vertical-align:middle;
					  font-weight:300;
					  font-size:12px;
					  text-shadow: -1px -1px 1px rgba(0, 0, 0, 0.1);
					  border-right: 1px solid #C1C3D1;
					}

					td:last-child {
					  border-right: 0px;
					}

					th.text-left {
					  text-align: left;
					}

					th.text-center {
					  text-align: center;
					}

					th.text-right {
					  text-align: right;
					}

					td.text-left {
					  text-align: left;
					}

					td.text-center {
					  text-align: center;
					}

					td.text-right {
					  text-align: right;
					}
					"
			}
            FlorAntara{
                $Css= "*{
                    box-sizing: border-box;
                    -webkit-box-sizing: border-box;
                    -moz-box-sizing: border-box;
                }
                body{
                    font-family: Helvetica;
                    -webkit-font-smoothing: antialiased;
                    background: rgba( 71, 147, 227, 1);
                }
                h2{
                    text-align: center;
                    font-size: 18px;
                    text-transform: uppercase;
                    letter-spacing: 1px;
                    color: white;
                    padding: 30px 0;
                }
                
                /* Table Styles */
                
                .table-wrapper{
                    margin: 10px 70px 70px;
                    box-shadow: 0px 35px 50px rgba( 0, 0, 0, 0.2 );
                }
                
                .fl-table {
                    border-radius: 5px;
                    font-size: 12px;
                    font-weight: normal;
                    border: none;
                    border-collapse: collapse;
                    width: 100%;
                    max-width: 100%;
                    white-space: nowrap;
                    background-color: white;
                }
                
                .fl-table td, .fl-table th {
                    text-align: center;
                    padding: 8px;
                }
                
                .fl-table td {
                    border-right: 1px solid #f8f8f8;
                    font-size: 12px;
                }
                
                .fl-table thead th {
                    color: #ffffff;
                    background: #4FC3A1;
                }
                
                
                .fl-table thead th:nth-child(odd) {
                    color: #ffffff;
                    background: #324960;
                }
                
                .fl-table tr:nth-child(even) {
                    background: #F8F8F8;
                }
                
                /* Responsive */
                
                @media (max-width: 767px) {
                    .fl-table {
                        display: block;
                        width: 100%;
                    }
                    .table-wrapper:before{
                        content: 'Scroll horizontally >';
                        display: block;
                        text-align: right;
                        font-size: 11px;
                        color: white;
                        padding: 0 0 10px;
                    }
                    .fl-table thead, .fl-table tbody, .fl-table thead th {
                        display: block;
                    }
                    .fl-table thead th:last-child{
                        border-bottom: none;
                    }
                    .fl-table thead {
                        float: left;
                    }
                    .fl-table tbody {
                        width: auto;
                        position: relative;
                        overflow-x: auto;
                    }
                    .fl-table td, .fl-table th {
                        padding: 20px .625em .625em .625em;
                        height: 60px;
                        vertical-align: middle;
                        box-sizing: border-box;
                        overflow-x: hidden;
                        overflow-y: auto;
                        width: 120px;
                        font-size: 13px;
                        text-overflow: ellipsis;
                    }
                    .fl-table thead th {
                        text-align: left;
                        border-bottom: 1px solid #f7f7f9;
                    }
                    .fl-table tbody tr {
                        display: table-cell;
                    }
                    .fl-table tbody tr:nth-child(odd) {
                        background: none;
                    }
                    .fl-table tr:nth-child(even) {
                        background: transparent;
                    }
                    .fl-table tr td:nth-child(odd) {
                        background: #F8F8F8;
                        border-right: 1px solid #E6E4E4;
                    }
                    .fl-table tr td:nth-child(even) {
                        border-right: 1px solid #E6E4E4;
                    }
                    .fl-table tbody td {
                        display: block;
                        text-align: center;
                    }
                }"
            }
		}
		if ($PSBoundParameters['Verbose']) {RW-Logme -LogFileName $LogFilePath -text "RW-CssStyle: Completed"}
		return $CSS
	}
}

$APIKey = # you need register an account for CoinMarket API @ https://coinmarketcap.com/api/
$URL = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest'
$CsvPath = "$(split-path -parent $MyInvocation.MyCommand.Definition)\$(get-date -Format MM-dd-yyyy).csv"

$SMTP ='sendgrid.net'
$SMTPUser= '' # you will need to register an account on sendgrid.net for API 
$SMTPPwd = '' # you will need to register an account on sendgrid.net for API password
$Sender='' #email sender address 
$Reciever='' 
$Subject = 'CoinMarket Report'

$Report = GetReport -Data $(GetData -URL $URL -APIKey $APIKey).data
$Report | Export-Csv -Path $CsvPath -NoTypeInformation
$HTML = $Report | ConvertTo-Html -Head "<style>$(CssStyle -Style Table )</style>" | Out-String
Send-MailMessage -From $Sender -BodyAsHtml -Body $HTML -Subject $Subject -to $Reciever -SmtpServer $SMTP -Credential (CredentialObj -UserName $SMTPUser -Password $SMTPPwd) -Attachments $CsvPath
