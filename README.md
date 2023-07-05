# VBA-Lichess-Data-Aanlyse

# Omschrijving:
Voor de opkomst van online schaakservers waar je gratis online kunt schaken (denk aan Chess.com en Lichess.org), waren er twee commerciele schaakservers die erg populair waren: ICC en Playchess. 
Playchess is een online schaakserver die is ontwikkeld en wordt beheerd door het Duitse bedrijf Chessbase. Er was een tijd dat als je een schaakprogramma van Chessbase kocht (denk bijvoorbeeld aan een versie van het schaakprogramma ‘Fritz’), dan kreeg je hierbij automatisch ook toegang tot Playchess.com. Alle partijen die je op Playchess speelde, werden automatisch opgeslagen in een database. Deze database kun je dan vervolgens via je Chessbase-schaakprogramma openen en daarmee de gespeelde partijen opnieuw afspelen. 
Het Chessbase-schaakprogramma biedt ook de mogelijkheid om de complete database met schaakpartijen te exporteren naar een tekstbestand. Een screenshot van hoe dit tekstbestand eruit ziet: 

![image](https://github.com/priksten/VBA-Lichess-Data-Aanlyse/assets/85739742/f07200a7-2f23-472c-89da-cc1fada17940)

Het doel van dit project was om met behulp van VBA in Excel een programma te schrjiven dat de schaker inzicht geeft in zijn prestaties. Hierbij gaat het om hoe de speler presteert:
•	Prestaties over de partijen: hoeveel winst, hoeveel remise, hoeveel verlies
•	Overzicht van het niveau van de tegenstanders
•	Het ratingverloop van de speler over deze partijen in een grafiek, waarbij ook de maximaal behaalde rating wordt aangegeven. 
•	De mogelijkheid bieden om de resultaten die met wit behaald zijn te vergelijken met de resultaten die met zwart behaald zijn

# Programmeertaal + versie: 
Microsoft 365 (Excel + VBA in Excel)

# Screenshots:

![image](https://github.com/priksten/VBA-Lichess-Data-Aanlyse/assets/85739742/ed5c79a7-7d3e-4b58-bbc6-7c545cbf4d5f)

De eerste stap in het programma is om de schaakpartijen vanuit het tekstbestand te importeren in Excel. Vervolgens zijn er een aantal macro’s nodig om uit de schaakpartijen alleen gegevens te halen zoals de datum, namen van de spelers, ratings van beide spelers en de uitslag van de partij en deze te tonen in de vorm van een lijst zoals in bovenstaande screenshot. Vanwege onregelmatigheden in de manier waarop de gegevens opgeslagen zijn, zijn er een aantal extra stappen nodig bepaalde correcties in de lijst aan te brengen. 

![image](https://github.com/priksten/VBA-Lichess-Data-Aanlyse/assets/85739742/38a60649-c155-485f-9bba-c5465d29510f)

Het programma biedt de mogelijkheid om de behaalde resultaten in verschillende spelvormen (blitz, bullet en rapid) te analyseren. In het voorbeeld is te zien hoe stap voor stap de Blitz-resultaten geanalyseerd worden. Allereerst worden alle Blitz-partijen uit de lijst van het vorige screenshot gehaald en in een lijst weergegeven. Maar deze lijst moet nog omgewerkt worden voor het maken van de analyse:

![image](https://github.com/priksten/VBA-Lichess-Data-Aanlyse/assets/85739742/1571b3bb-24ef-4c70-b808-40a8eb78ef47)

In deze lijst staan de gegevens waarop de analyse gebaseerd is: nummer partij, rating speler, rating tegenstander en resultaat. 

![image](https://github.com/priksten/VBA-Lichess-Data-Aanlyse/assets/85739742/3768b9c7-15da-49bb-8ce5-eaec8f448d53)

Het eindresultaat van de analyse. 
