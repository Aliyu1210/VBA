{\rtf1\ansi\ansicpg1252\cocoartf2639
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fnil\fcharset0 Menlo-Regular;}
{\colortbl;\red255\green255\blue255;\red70\green137\blue204;\red23\green23\blue23;\red202\green202\blue202;
\red89\green138\blue67;\red194\green126\blue101;\red167\green197\blue152;}
{\*\expandedcolortbl;;\cssrgb\c33725\c61176\c83922;\cssrgb\c11765\c11765\c11765;\cssrgb\c83137\c83137\c83137;
\cssrgb\c41569\c60000\c33333;\cssrgb\c80784\c56863\c47059;\cssrgb\c70980\c80784\c65882;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs24 \cf2 \cb3 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 Sub\cf4 \strokec4  Stock_Data():\cb1 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 'Defining Headers\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3     Range(\cf6 \strokec6 "i1"\cf4 \strokec4 ).Value = \cf6 \strokec6 "Ticker"\cf4 \cb1 \strokec4 \
\cb3     Range(\cf6 \strokec6 "j1"\cf4 \strokec4 ).Value = \cf6 \strokec6 "Yearly Change"\cf4 \cb1 \strokec4 \
\cb3     Range(\cf6 \strokec6 "k1"\cf4 \strokec4 ).Value = \cf6 \strokec6 "Percent change"\cf4 \cb1 \strokec4 \
\cb3     Range(\cf6 \strokec6 "l1"\cf4 \strokec4 ).Value = \cf6 \strokec6 "Total Stock Volume"\cf4 \cb1 \strokec4 \
\cb3     Range(\cf6 \strokec6 "p1"\cf4 \strokec4 ).Value = \cf6 \strokec6 "Ticker"\cf4 \cb1 \strokec4 \
\cb3     Range(\cf6 \strokec6 "q1"\cf4 \strokec4 ).Value = \cf6 \strokec6 "Value"\cf4 \cb1 \strokec4 \
\cb3     Range(\cf6 \strokec6 "o2"\cf4 \strokec4 ).Value = \cf6 \strokec6 "Greatest % increase"\cf4 \cb1 \strokec4 \
\cb3     Range(\cf6 \strokec6 "o3"\cf4 \strokec4 ).Value = \cf6 \strokec6 "Greatest % Decrease"\cf4 \cb1 \strokec4 \
\cb3     Range(\cf6 \strokec6 "o4"\cf4 \strokec4 ).Value = \cf6 \strokec6 " Greatest Total Volume"\cf4 \cb1 \strokec4 \
\cb3     \cf5 \strokec5 'Declaring variale types\cf4 \cb1 \strokec4 \
\cb3     \cf2 \strokec2 Dim\cf4 \strokec4  Ticker As String\cb1 \
\cb3     \cf5 \strokec5 'Dim YearlyChange As Long\cf4 \cb1 \strokec4 \
\cb3     \cf5 \strokec5 'Dim total As Long\cf4 \cb1 \strokec4 \
\cb3     \cf2 \strokec2 Dim\cf4 \strokec4  Openprice As Long\cb1 \
\cb3     \cf2 \strokec2 Dim\cf4 \strokec4  Closing As Double\cb1 \
\cb3     \cf2 \strokec2 Dim\cf4 \strokec4  percentchange As Double\cb1 \
\cb3     \cf2 \strokec2 Dim\cf4 \strokec4  r As Long\cb1 \
\cb3     \cf2 \strokec2 Dim\cf4 \strokec4  i  As Long\cb1 \
\cb3     \cb1 \
\cb3     Openprice = Range(\cf6 \strokec6 "c2"\cf4 \strokec4 ).Value\cb1 \
\cb3     percentchange = \cf7 \strokec7 0\cf4 \cb1 \strokec4 \
\cb3    \cf2 \strokec2 For\cf4 \strokec4  \cf2 \strokec2 Each\cf4 \strokec4  ws In Worksheets\cb1 \
\cb3     RowCount = Cells( Rows.Count, \cf6 \strokec6 "A"\cf4 \strokec4 ).End(xlUp).Row\cb1 \
\cb3     \cb1 \
\cb3     \cf2 \strokec2 Dim\cf4 \strokec4  j As Integer\cb1 \
\cb3     j = \cf7 \strokec7 0\cf4 \cb1 \strokec4 \
\cb3     \cf2 \strokec2 For\cf4 \strokec4  i = \cf7 \strokec7 2\cf4 \strokec4  \cf2 \strokec2 To\cf4 \strokec4  RowCount\cb1 \
\cb3     \cb1 \
\cb3      Total = Total + ws.Cells(i, 7).Value\cb1 \
\cb3     \cf2 \strokec2 If\cf4 \strokec4  ws.Cells(i + \cf7 \strokec7 1\cf4 \strokec4 , \cf6 \strokec6 "A"\cf4 \strokec4 ).Value <> ws.Cells(i, \cf6 \strokec6 "A"\cf4 \strokec4 ).Value \cf2 \strokec2 Then\cf4 \cb1 \strokec4 \
\cb3     Ticker = ws.Cells(i, \cf6 \strokec6 "A"\cf4 \strokec4 ).Value\cb1 \
\cb3     Closing = ws.Cells(i, \cf6 \strokec6 "f"\cf4 \strokec4 ).Value\cb1 \
\cb3     \cf2 \strokec2 Else\cf4 \cb1 \strokec4 \
\cb3     YearlyChange = Closing - Openprice\cb1 \
\cb3     \cf2 \strokec2 End if\cf4 \cb1 \strokec4 \
\cb3     \cf2 \strokec2 If\cf4 \strokec4  Openprice = \cf7 \strokec7 0\cf4 \strokec4  \cf2 \strokec2 Then\cf4 \cb1 \strokec4 \
\cb3     percentchange = \cf7 \strokec7 0\cf4 \cb1 \strokec4 \
\cb3     \cf2 \strokec2 Else\cf4 \cb1 \strokec4 \
\cb3     \cb1 \
\cb3     percentchange = YearlyChange / Openprice\cb1 \
\cb3     \cf2 \strokec2 End If\cf4 \cb1 \strokec4 \
\cb3     \cb1 \
\cb3     \cb1 \
\cb3     ws.Cells(\cf7 \strokec7 2\cf4 \strokec4  + j, \cf6 \strokec6 "I"\cf4 \strokec4 ).Value = Ticker\cb1 \
\cb3     ws.Cells(\cf7 \strokec7 2\cf4 \strokec4  + j, \cf6 \strokec6 "j"\cf4 \strokec4 ).Value = YearlyChange\cb1 \
\cb3     ws.Cells(\cf7 \strokec7 2\cf4 \strokec4  + j, \cf6 \strokec6 "K"\cf4 \strokec4 ).Value = percentchange\cb1 \
\cb3     \cf5 \strokec5 'Range("I" & 2 + j).Value = Ticker\cf4 \cb1 \strokec4 \
\cb3     \cf5 \strokec5 'Range("j" & RowCount).Value = YearlyChange\cf4 \cb1 \strokec4 \
\cb3    \cf5 \strokec5 'conditional formating\cf4 \cb1 \strokec4 \
\cb3    \cf2 \strokec2 If\cf4 \strokec4  ws.Cells(\cf7 \strokec7 2\cf4 \strokec4  + j, \cf6 \strokec6 "J"\cf4 \strokec4 ).Value > \cf7 \strokec7 0\cf4 \strokec4  \cf2 \strokec2 Then\cf4 \cb1 \strokec4 \
\cb3         ws.Cells(\cf7 \strokec7 2\cf4 \strokec4  + j, \cf6 \strokec6 "J"\cf4 \strokec4 ).Interior.ColorIndex = \cf7 \strokec7 4\cf4 \cb1 \strokec4 \
\cb3         \cf2 \strokec2 Else\cf4 \cb1 \strokec4 \
\cb3         ws.Cells(\cf7 \strokec7 2\cf4 \strokec4  + j, \cf6 \strokec6 "J"\cf4 \strokec4 ).Interior.ColorIndex = \cf7 \strokec7 3\cf4 \cb1 \strokec4 \
\cb3     \cf2 \strokec2 End If\cf4 \cb1 \strokec4 \
\cb3         \cf2 \strokec2 For\cf4 \strokec4  r = \cf7 \strokec7 2\cf4 \strokec4  \cf2 \strokec2 To\cf4 \strokec4  RowCount\cb1 \
\cb3         Greatest_Increase = ws.Cells(i, \cf6 \strokec6 "k"\cf4 \strokec4 ).Value\cb1 \
\cb3         Greatest_Decrease = ws.Cells(i, \cf6 \strokec6 "K"\cf4 \strokec4 ).Value\cb1 \
\cb3         Greatest_Total = ws.Cells(i, \cf6 \strokec6 "L"\cf4 \strokec4 ).Value\cb1 \
\cb3          \cf2 \strokec2 If\cf4 \strokec4  ws.Range(\cf6 \strokec6 "K"\cf4 \strokec4  & r + \cf7 \strokec7 1\cf4 \strokec4 ).Value > Greatest_Increase \cf2 \strokec2 Then\cf4 \cb1 \strokec4 \
\cb3             Greatest_Increase = ws.Range(\cf6 \strokec6 "K"\cf4 \strokec4  & r + \cf7 \strokec7 1\cf4 \strokec4 ).Value\cb1 \
\cb3             Greatest_Increase_Ticker = ws.Range(\cf6 \strokec6 "I"\cf4 \strokec4  & r + \cf7 \strokec7 1\cf4 \strokec4 ).Value\cb1 \
\cb3             \cf2 \strokec2 ElseIf\cf4 \strokec4  ws.Range(\cf6 \strokec6 "K"\cf4 \strokec4  & r + \cf7 \strokec7 1\cf4 \strokec4 ).Value < Greatest_Decrease \cf2 \strokec2 Then\cf4 \cb1 \strokec4 \
\cb3              Greatest_Total = ws.Range(\cf6 \strokec6 "L"\cf4 \strokec4  & r + \cf7 \strokec7 1\cf4 \strokec4 ).Value\cb1 \
\cb3              Greatest_Total_Ticker = ws.Range(\cf6 \strokec6 "I"\cf4 \strokec4  & r + \cf7 \strokec7 1\cf4 \strokec4 ).Value\cb1 \
\cb3              \cf5 \strokec5 'total = total + Cells(i, 7).Value\cf4 \cb1 \strokec4 \
\cb3               ws.Range(\cf6 \strokec6 "L"\cf4 \strokec4  & \cf7 \strokec7 2\cf4 \strokec4  + j).Value = Total\cb1 \
\cb3              Openprice = ws.Cells(i + \cf7 \strokec7 1\cf4 \strokec4 , \cf6 \strokec6 "C"\cf4 \strokec4 ).Value\cb1 \
\cb3     \cb1 \
\cb3              Total = \cf7 \strokec7 0\cf4 \cb1 \strokec4 \
\cb3     \cb1 \
\cb3              j = j + \cf7 \strokec7 1\cf4 \cb1 \strokec4 \
\cb3     \cb1 \
\cb3     \cb1 \
\cb3     \cf2 \strokec2 End If\cf4 \cb1 \strokec4 \
\cb3    \cb1 \
\cb3     \cf2 \strokec2 Next\cf4 \strokec4  r\cb1 \
\cb3     \cf2 \strokec2 Next\cf4 \strokec4  i\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 Next\cf4 \strokec4  ws\cb1 \
\
\cf2 \cb3 \strokec2 End Sub\cf4 \cb1 \strokec4 \
}