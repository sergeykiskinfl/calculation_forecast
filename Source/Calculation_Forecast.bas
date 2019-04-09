'  Copyright 2019 Sergey Kiskin

'  Licensed under the Apache License, Version 2.0 (the "License");
'  you may not use this file except in compliance with the License.
'  You may obtain a copy of the License at
'
'      http://www.apache.org/licenses/LICENSE-2.0
'
'  Unless required by applicable law or agreed to in writing, software
'  distributed under the License is distributed on an "AS IS" BASIS,
'  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or  implied.
'  See the License for the specific language governing permissions and
'  limitations under the License.

Option Explicit
Dim EA As Excel.Application
Dim WB As Excel.Workbook
Dim WS(0 To 1) As Excel.Worksheet
Dim i&, y&
Dim lvalue&
Dim nameField(0 To 6) 'Название заголовков
Dim iLastRow&(0 To 1) 'Массив номеров последних строк на листах
Dim iLastColumn&(0 To 10) 'Массив номеров последних столбцов на листах
Dim WSR(0 To 1) As Range 'Массив для диапазонов
Dim nPeriod& 'Период для планирования
Dim nRow$ 'Номер суммируемой строки
Dim nColumn&(0 To 6) 'Номера столбцов на листе "Прогноз"
Dim dTotal# 'Сумма продаж
Dim dTotalMonth# 'Сумма продаж за одинаковые месяцы
Dim dSeasonFactor#() 'Коэффициент сезонности
Dim dForecast#() 'Прогноз
Dim nLastRowInData& 'Последняя строка с первоначальными данными
Dim dValueX#, dAlfa#
Dim dStDev# 'Стандартное отклонение
Dim dConf# 'Допустимое стандартное отклонение
Dim CH As Object 'График прогноза

Public Sub calculationForecast()

Set EA = Excel.Application

With EA

    .ScreenUpdating = False: .DisplayAlerts = False: .StatusBar = False

End With

'Задание основных переменных
nameField(0) = "Период"
nameField(1) = "Продажи, тыс.руб."
nameField(2) = "Прогноз"
nameField(3) = "Оптимистичный"
nameField(4) = "Пессимистичный"
nameField(5) = "Коэффициент сезонности"
nameField(6) = "Отклонение"

nColumn(0) = 1
nColumn(1) = 2

nPeriod = 12
dAlfa = 0.05

ReDim dSeasonFactor(0 To nPeriod - 1)
ReDim dForecast(0 To nPeriod - 1)

On Error Resume Next

Set WB = EA.Workbooks("Calculation_Forecast.xlsm")

If WB Is Nothing Then Set WB = EA.ActiveWorkbook

Set WS(0) = WB.Worksheets("Исходные данные")
Set WS(1) = WB.Worksheets("Прогноз")

'Удаление устаревших данных, формирование листа "Прогноз"
If WS(1) Is Nothing Then

    WB.Sheets.Add(, Sheets(Sheets.Count)).Name = "Прогноз"

Else:

    WS(1).Delete

    WB.Sheets.Add(, Sheets(Sheets.Count)).Name = "Прогноз"

End If

Set WS(1) = WB.Worksheets("Прогноз")

iLastRow(0) = Utils_.lastRow(WS(0))
iLastColumn(0) = Utils_.lastColumn(WS(0))

'Копирование данных с листа "Исходные данные" на лист "Прогноз"
With WS(0)

    Set WSR(0) = .Range(.Cells(1, 1), .Cells(iLastRow(0), iLastColumn(0)))

End With

WSR(0).Copy WS(1).Cells(1, 1)

iLastRow(1) = Utils_.lastRow(WS(1))
iLastColumn(1) = Utils_.lastColumn(WS(1))

With WS(1)

    For i = 2 To UBound(nameField)
          
       .Cells(1, i + 1) = nameField(i)
       nColumn(i) = i + 1
        
    Next i
    
    lvalue = .Cells(iLastRow(1), 2).value
    
    For i = 3 To 5
      
        .Cells(iLastRow(1), i).value = lvalue

    Next i
    
    iLastRow(1) = Utils_.lastRow(WS(1))
    iLastColumn(1) = Utils_.lastColumn(WS(1))
    
    nLastRowInData = iLastRow(1)
    
    'Расчет коэффициента сезонности
    Set WSR(0) = .Range(.Cells(2, nColumn(1)), .Cells(iLastRow(1), nColumn(1)))
    
    dTotal = EA.WorksheetFunction.Sum(WSR(0))
              
    For i = 2 To nPeriod + 1
              
        nRow = i
        dTotalMonth = 0
        
        For y = 0 To 2
            
            dTotalMonth = dTotalMonth + .Cells(nRow, nColumn(1)).value
            nRow = nRow + nPeriod
        
        Next y
        
        dSeasonFactor(i - 2) = (dTotalMonth / dTotal) * 12
        
        .Cells(i, nColumn(5)).value = dSeasonFactor(i - 2)
        
    Next i
    
    Set WSR(0) = .Range(.Cells(2, nColumn(5)), .Cells(nPeriod + 1, nColumn(5)))
    
    WSR(0).NumberFormat = "0.00%"
    
    'Ввод дополнительного периода для прогнозирования
    Set WSR(0) = .Range(.Cells(iLastRow(1) - 1, nColumn(0)), .Cells(iLastRow(1), nColumn(0)))
    Set WSR(1) = .Range(.Cells(iLastRow(1) - 1, nColumn(0)), .Cells(iLastRow(1) + nPeriod, nColumn(0)))
    WSR(0).AutoFill WSR(1)
    
    iLastRow(1) = Utils_.lastRow(WS(1))
    
    nRow = 0
    
    For i = nLastRowInData + 1 To iLastRow(1)
          
        dValueX = .Cells(i, nColumn(0)).value
        'Debug.Print (dValueX)
        Set WSR(0) = .Range(.Cells(2, nColumn(0)), .Cells(nLastRowInData, nColumn(0)))
        Set WSR(1) = .Range(.Cells(2, nColumn(1)), .Cells(nLastRowInData, nColumn(1)))
        
        'Расчет прогноза
        dForecast(nRow) = EA.WorksheetFunction.Forecast(dValueX, WSR(1), WSR(0))
        'Debug.Print (dForecast(nRow))
        
        'Поправка на коэффициент сезонности
        .Cells(i, nColumn(2)).value = EA.WorksheetFunction.Round(dForecast(nRow) * dSeasonFactor(nRow), 2)
        
        nRow = nRow + 1
    
    Next i
    
    'Расчет допустимого стандартного отклонения
    Set WSR(0) = .Range(.Cells(nLastRowInData + 1, nColumn(2)), .Cells(iLastRow(1), nColumn(2)))
        
    With EA.WorksheetFunction
    
        dStDev = .StDev(WSR(0))
        dConf = .Round(.Confidence(dAlfa, dStDev, nPeriod), 2)
    
    End With
    
    .Cells(2, nColumn(6)).value = dConf
    
    nRow = 0
    
    For i = nLastRowInData + 1 To iLastRow(1)
                  
        .Cells(i, nColumn(3)).value = EA.WorksheetFunction.Round(.Cells(i, nColumn(2)).value + dConf, 2)
        .Cells(i, nColumn(4)).value = EA.WorksheetFunction.Round(.Cells(i, nColumn(2)).value - dConf, 2)
    
    Next i
     
    'Оформление шапки таблицы
    Set WSR(0) = .Range(.Cells(1, 1), .Cells(1, iLastColumn(1)))
    
    With WSR(0)

        .Interior.ThemeColor = xlThemeColorAccent6
        .Font.ThemeColor = xlThemeColorDark1
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .EntireColumn.AutoFit
        .Font.Bold = True
        .Borders.ThemeColor = xlThemeColorDark1
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlInsideVertical).Weight = xlThin

    End With
    
    'Построение графика
    Set WSR(0) = .Range(.Cells(1, 1), .Cells(iLastRow(1), nColumn(4)))
    Set CH = .ChartObjects.Add(400, 100, 500, 250)
    
    With CH.Chart
        
        .ChartWizard WSR(0), xlLine, Title:="Продажи"
        .SetElement msoElementLineDropLine
            
    End With
    
End With
    

With EA

    .ScreenUpdating = True: .DisplayAlerts = True

End With

End Sub



