# Analiza-razdelilnikov
Analiza .xls razdelilnikov od mojupravnik.si

Analizator za razdelilnike s portala mojUpravnik. 				Vid Mlačnik, 2025

Uporaba:
1.	Vpiši se v portal mojupravnik.si. 
2.	Odpri podrobnosti delilnika za nek mesec.

  	![image](https://github.com/user-attachments/assets/f2cfcf70-7e9b-45f4-b8b0-15bb2494b17c)
4.	Na dnu strani izvozi XLS.  
    ![image](https://github.com/user-attachments/assets/b1b235eb-a131-48e9-8e35-e9668734ab9f)
5.	Ponovi ta izvoz za vse mesce v Arhivu razdelilnikov
6.	Prenesene razdelilnike dodaj v mapo k programu. Program bo prebral le datoteke v isti mapi z začetnico "Razdelilnik_".

    ![image](https://github.com/user-attachments/assets/67b3d1d8-31ac-4f7f-a42a-5bb8b0618ea5)
8.	Zaženi program. 
Po nekaj trenutkih bo program v brskalniku prikazal grafično analizo podatkov. 
V mapi programa se pojavi "Rezultati za vse razdelilnike.xlsx" z urejenimi podatki.




Funkcije:
Grafični povzetek 
![image](https://github.com/user-attachments/assets/7a29ac65-d80e-4b6b-b35c-f2c5ba218de3)

Excel tabele "Rezultati za vse razdelilnike.xlsx"
-	Tabela 'Povzetek' za vse stroške povzema mesečno povprečje, skupno vsoto vseh položnic, povprečni delež zneska na vaši položnici, skupno vsoto vseh položnic stavbe (glej (ERROR!) opozorilo spodaj) in preračunan delež stroška stanovanja glede na stavbo. 
-	Tabela 'Poraba' podaja porabo vode in preračunano razmerje stroška ogrevanja glede na stavbo.
-	Tabela 'Mesečni stroški stanovanja' je urejen povzetek vseh postavk na položnicah.
-	Tabela 'Mesečni stroški stavbe' je urejen povzetek vseh postavk na položnicah. Pozor, te niso nujno pravilno prepisane (glej (ERROR!) opozorilo spodaj).

[(ERROR!)]: Pozor – podatki skupnih stroškov stavbe so nesigurni, ker so navedene cene računov z mojUpravnik pri nekaterih postavkah ponovljene, pri nekaterih razdeljene. Program za tabelo "Mesečni stroški stavbe (ERROR!)" uporablja samo eno vrednost na postavko (zadnjo po vrsti).






Če ne dela (Troubleshooting): 
-	Preveri, da excel tabela ni odprta. 
-	Preveri, da je nastavljen privzeti brskalnik (web browser).
