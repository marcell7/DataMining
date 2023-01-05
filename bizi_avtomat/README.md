# bizi-avtomat
Python skripta, ki avtomatizira iskanje po spletni strani Bizi.si.

Delovanje je preprosto; Potrebuješ .xlsx datoteko z iskalnimi nizi. To so lahko davčne številke podjetij. Skripta potem uporabi vsak iskalni niz in poišče iskano podjetje. Nato pridobi podatke, ki so vidni na profilni strani posameznega podjetja. Na koncu se vsi podatki izvozijo v novo .xlsx datoteko.

- Uporabljeni sta dve nestandardni knjižnici, ki jih je potrebno namestiti -> Selenium (avtomatizacija brskalnika) in Pandas (branje in izvoz .xlsx datotek)
- Potrebuješ Webdriver (.exe datoteko dobiš na: https://github.com/mozilla/geckodriver/releases)
- V isti mapi, kjer je skripta naredi novo mapo "driver", v katero kopiraš prenešeno .exe datoteko
- Skripto poženeš. Ob zagonu je treba določiti še 3 argumente -> username, password in pot do vhodne .xlsx datoteke



