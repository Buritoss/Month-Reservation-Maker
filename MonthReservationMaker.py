import xlsxwriter
import datetime
import calendar

class MonthVerivicationWriter():
    def __init__(self):
        self.currentRow=0
        self.CurrentColumn=0
        
        self.date = datetime.datetime.today().replace(day=1)
        self.setDayToWorkDay()    

        self.workbook = xlsxwriter.Workbook(f'Weryfikacje druku {self.date.month}.{self.date.year}.xlsx') 
        self.worksheet = self.workbook.add_worksheet()

        self.Start()
        self.MakeOneMonth()

    def setDayToWorkDay(self):
        if self.date.weekday() == 5 or self.date.weekday()==6:
            self.date+=datetime.timedelta(7-self.date.weekday)

    def Start(self):
        self.worksheet.set_column("A:G",25)
        
        state = ["CZEKA","DRUKUJE SIĘ","DOSTARCZONE","WYDRUKOWANE"]
        self.worksheet.write_column("I1",state)

        place=["BIURO","USŁUGA ZEWNĘTRZNA","DRUKARKA/CZEKOLADA","DŁUGOPISY 3D","GABRYŚ","TECHDOKTOR","KRZYSIA","OKULARY VR"]
        self.worksheet.write_column("J1",place)
  

    def MakeOneMonth(self):
        lastDayOfTheMonth = calendar.monthrange(self.date.year,self.date.month)[1]
        self.writeTitles()
        while self.date.month == datetime.datetime.today().month:
            self.writeDay()

        self.workbook.close()

    def writeTitles(self):
        titles = ["Dzień Tygodnia","Miejsce","Typ wydruku","Ilość wydruku","Miejsce wydruku","Status wydruku","Informacje dodatkowe"]

        my_format = self.workbook.add_format()
        my_format.set_align('center')
        my_format.set_bg_color("yellow")
        my_format.set_border(1)

        for t in titles:
            self.worksheet.write(self.currentRow,self.CurrentColumn,t,my_format)
            self.CurrentColumn+=1
        
        self.currentRow+=1
        self.CurrentColumn=0

    def writeDay(self):
        merge_format = self.workbook.add_format()
        merge_format.set_align("center")
        merge_format.set_align("vcenter")
        merge_format.set_bg_color("green")
        merge_format.set_border(1)

        dayName=""

        match self.date.weekday():
            case 0:
                dayName="Poniedziałek"
            case 1:
                dayName="Wtorek"
            case 2:
                dayName="Środa"
            case 3:
                dayName="Czwartek"
            case 4:
                dayName="Piątek"

        self.worksheet.merge_range(self.currentRow,
                                    self.CurrentColumn,
                                    self.currentRow+4,
                                    self.CurrentColumn,
                                    f"{dayName} {self.date.strftime("%d-%m-%Y")}",
                                    merge_format)
        self.writeLists()

        self.currentRow+=5
        self.CurrentColumn=0
        if self.date.weekday()==4:
            self.date+=datetime.timedelta(days=3)
            self.writeTitles()
        else:
            self.date+=datetime.timedelta(days=1)

    def writeLists(self):
        self.CurrentColumn+=4
        self.worksheet.data_validation(
            self.currentRow,
            self.CurrentColumn,
            self.currentRow+4,
            self.CurrentColumn,
            {
                'validate': 'list',
                'source': "J:J",
            }
        )

        self.CurrentColumn+=1
        self.worksheet.data_validation(
            self.currentRow,
            self.CurrentColumn,
            self.currentRow+4,
            self.CurrentColumn,
            {
                'validate': 'list',
                'source': "I:I",
            }
        )

 