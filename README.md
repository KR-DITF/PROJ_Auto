# PROJ_Auto
Projekts Lietojumprogrammas automatizēšanas kursā

Uzdevums.
Faila “saraksts_IMEI_konti.xls” tabā Konti ir klientu saraksts, kuriem ir veikta pirmreizējā
apstrāde. Vienu reizi mēnesī šis saraksts tiek papildināts ar jauniem klientiem. Klienti
neatkārtojas. Klientu numurs ir norādīts kolonnā account Nr. Pievienojot klientus šim
sarakstam, failā attiecīgi tiek pievienots jauns tabs ar klienta iekārtu numuriem. Katrs klients
var būt reģistrēts tikai vienā tabā. Vienam klientam attiecīgajā tabā var būt vairākas iekārtas.
Regulāri tiek veikta klientu otrreizējā apstrāde, kuras rezultātā klientam var tikt aizpildītas
viena vai vairākas ailes (galvenā otrreizēji apstrādātā klienta pazīme – ar burtu aizpildīta
kolonna Statuss).
Ir pieejams papildus fails “aktualie.xls”, kurā ir klienti, kuru aktuālā bilance ir vai nu parāds
(pozitīvs skaitlis kolonnā TOTAL_BALANCE) vai nu pārmaksa (negatīvs skaitlis kolonnā
TOTAL_BALANCE). Ja klienta aktuālā bilance ir “0”, tad viņa šajā failā nav.
Uzdevums: saņemot jaunu failu “aktualie.xls”, veikt apstrādi otrreizēji neapstrādātajiem
klientiem no faila “saraksts_IMEI_konti.xls”, pārbaudot klienta bilanci – ja tā ir mazāka par
10, tad no faila tabiem jāatrod attiecīgā klienta visi iekārtu numuri.
Rezultātā jāizdod tabula ar šādām kolonnām:
Konta nr., MSISDN, IMEI, IMEI2
Tabulu var ievietot vai nu faila “saraksts_IMEI_konti.xls” jaunā tabā vai arī nosūtīt epastā.

Izmantotās bibliotēkas

openpyxl - Izmantots lai atvērtu, aizvērtu un varētu izmantot excel lapas un datus

Programmatūra izveidota specifiskai darbībai, lai automatizētu daļu no personas darba, tapēc nevar tikt pielietota dažādi. Rakstīta specifiskām prasībām un lielumiem, un failiem
