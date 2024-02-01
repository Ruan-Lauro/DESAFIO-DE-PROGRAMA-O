import gspread
import re


CODE = '1Mx7dL15850KsZirob7D37QfIlSR4KPrZwToi6MDAEkA'

gc = gspread.service_account(filename = 'key.json')
sh = gc.open_by_key(CODE)

ws = sh.worksheet('engenharia_de_software')


values = ws.get_all_values()

Total_absences = values[1:2]
Total_absences = Total_absences[0]
Total_absences = Total_absences[0]
Total_absences = re.findall(r'\d+', Total_absences)
Total_absences = int(Total_absences[0])
Total_absences = Total_absences*0.25

selected_rows = values[3:27]  

i = 1

for row in selected_rows:
    name = row[1:2]
    name = name[0]
    absence = row[2:3]
    p1 = row[3:4]
    p2 = row[4:5]
    p3 = row[5:6]
    
    absence = int(absence[0])
    Average = round(((int(p1[0])/10) + (int(p2[0])/10) + (int(p3[0])/10))/3, 1)

    if(absence > Total_absences):
        Situation = "Reprovado por Falta"
    else:
        if(Average < 5):
            Situation = "Reprovado"
        elif(5 <= Average < 7):
            Situation = "Exame Final"
        elif(Average >= 7):
            Situation = "Aprovado"
    
    # 5 <= (m + naf)/2
    # 10 <= m + naf
    # 10 − m <= naf
    # We can say that: naf = 10 - m  or naf > 10 - m
    # We will use this: naf = 10 - m

    if(Situation == "Exame Final"):
        naf = round(10 - Average,1)
    else:
        naf = 0

    print(f"Name: {name}, Absence: {absence}, test one: {int(p1[0])/10}, test two: {int(p2[0])/10}, test three: {int(p3[0])/10}, Average: {Average}, Situation: {Situation}, NAF: {naf}")
    print()

    ws.update_cell(i + 3, 7, Situation )
    ws.update_cell(i + 3, 8, naf )

    i += 1

    

    
    