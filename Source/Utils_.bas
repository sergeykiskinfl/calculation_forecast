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

Public Function lastColumn(WSLastColumnFunc As Excel.Worksheet, Optional ByVal iStartRowLastColumn& = 1)

'Функция возвращает номер последнего заполненного столбца определенной строки на листе.

With WSLastColumnFunc

lastColumn = .Cells(iStartRowLastColumn, .Columns.Count).End(xlToLeft).Column

End With

End Function

Public Function lastRow(WSLastRowFunc As Excel.Worksheet, Optional ByVal iStartColumnLastRow& = 1)

'Функция возвращает номер последней заполненной строки определенного столбца на листе.

With WSLastRowFunc
 
lastRow = .Cells(.Rows.Count, iStartColumnLastRow).End(xlUp).Row

End With

End Function


