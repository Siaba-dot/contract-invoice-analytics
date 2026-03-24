Sutarčių ir sąskaitų registras – Streamlit app
Ši Streamlit aplikacija skirta tavo Excel registrui analizuoti.
Ką rodo aplikacija
kiek šiuo metu yra galiojančių sutarčių
kiek yra neterminuotų sutarčių (`2100-12-31`)
kiek sutarčių baigiasi šį mėnesį
kaip kinta apkrova kas mėnesį
kiek sąskaitų yra `Neišrašyta` kiekvieną mėnesį
kurie klientai / sutartys yra `Neišrašyta` pasirinktą mėnesį
kiek sutarčių atėjo naujų ir kiek pasibaigė pagal mėnesius
kokie sezonai ar sutartys baigiasi artimiausiu metu
Kokio failo tikisi aplikacija
Lape `Registras` turi būti bent šie stulpeliai:
`Klientas`
`Sutarties Nr.`
`Galioja nuo`
`Galioja iki`
Taip pat turi būti mėnesių stulpeliai, pvz.:
`Rugsėjis`
`Spalis`
`Lapkritis`
...
`Sausis`
`Vasaris`
Statusai mėnesių stulpeliuose:
`Išrašyta`
`Neišrašyta`
Kaip įkelti į GitHub
Susikurk naują GitHub repository
Įkelk šiuos failus į repo šaknį:
`app.py`
`requirements.txt`
`runtime.txt`
`.gitignore`
`README.md`
Commit ir push
Kaip deployinti per Streamlit Cloud
Prisijunk prie Streamlit Cloud
Spausk New app
Pasirink savo GitHub repository
`Main file path` nurodyk: `app.py`
Spausk Deploy
Kaip naudoti
Į atidariusią aplikaciją, įkelk savo Excel failą
Pasirink duomenų lapą (pagal nutylėjimą `Registras`)
Šoninėje juostoje nurodyk pirmo mėnesio stulpelio metus
Peržiūrėk KPI, grafikus ir lenteles
Jei reikia – atsisiųsk analizės Excel
Pastabos
`2100-12-31` traktuojama kaip neterminuota sutartis
jei pirmas mėnuo faile yra `Rugsėjis`, pagal nutylėjimą starto metai nustatyti į `2024`
Galima vėliau pridėti:
filtrus pagal vadybininką
filtrus pagal įmonę
filtrus pagal pateikimo būdą
raudonus perspėjimus dėl artėjančių kainų / sezonų pakeitimų
TOP klientų, kur dažniausiai lieka neišrašyta, analizę
