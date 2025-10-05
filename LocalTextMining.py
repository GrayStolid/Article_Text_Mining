"""
Start date:
1400/06/31

End date:
1400/07/02

The purpose of this project is to find
similarities between two text columns
of data in .xlsx file format.
(Text Mining)
"""

##pandas module for open .xlsx file
##numpy for shape and reshape ndarray (or list)
##Time to calculate runtime
import pandas
import numpy
import time

##Get Input
def GetInp():
  while True:
      ##User Input
      UseInp = input("\nEnter the operation number:\n")
      try:
          UseInp = int(UseInp)
          if 0 <= UseInp and UseInp <= 6:
              return UseInp  
      except:
          pass
      print("\nPlease enter the operations number!\n")
      
def Prepr(Colum1, Colum2):
  Colum1 = list(Colum1)
  Colum2 = list(Colum2)
  if len(Colum1) == 1:
    Colum1 = [Colum1]
  if len(Colum2) == 1:
    Colum2 = [Colum2]
  for Tex1In in range(len(Colum1)):
    Colum1[Tex1In] = Colum1[Tex1In][0].split('; ')
    for Sub1In in range(len(Colum1[Tex1In])):
      Colum1[Tex1In][Sub1In] = ((Colum1[Tex1In][Sub1In]).lower()).replace('.','')
  for Tex2In in range(len(Colum2)):
    Colum2[Tex2In] = Colum2[Tex2In][0].split('; ')
    for Sub2In in range(len(Colum2[Tex2In])):
      Colum2[Tex2In][Sub2In] = ((Colum2[Tex2In][Sub2In]).lower()).replace('.','')
  return (Colum1, Colum2)

##Total Similarity of 1 by 1 Subjet
def TotaSimi():
  ##Text Mining Result, Total Similarity List
  global TextMiniResu, TotSimLis
  ##Total Similarity value
  TotalSimi = 0
  for i in range(len(TextMiniResu)):
      for j in range(numpy.shape(TextMiniResu[i])[0]):
          TotalSimi += TextMiniResu[i][j][0]
  TotSimLis.append(TotalSimi)

##Number and Location of same characters     
def NumLocSamChr(Colum1, Colum2, Ratio, UseInp):
  global TextMiniResu
  for i in range(len(Colum1)):
      Text1 = Colum1[i][0].split('; ')
      for j in range(len(Text1)):
          Subj1 = Text1[j]
          if Subj1 == 'nan':
              continue
          for k in range(len(Colum2)):
              Text2 = Colum2[k][0].split('; ')
              for l in range(len(Text2)):
                  Subj2 = Text2[l]
                  if Subj2 == 'nan':
                      continue
                  ##Slice Size of Text
                  ##Maximum Size of Text Slice
                  ##Number of fit Slice size
                  SliSiz = min(len(Subj1), len(Subj2))
                  MaxSli = max(len(Subj1), len(Subj2))
                  FitSli = MaxSli
                  LisRes = []
                  if UseInp != 0:
                    FitSli = 0
                    for x in range(MaxSli,MaxSli-SliSiz,-1):
                        FitSli += x
                  for m in range(1,SliSiz+1):
                      for n in range(len(Subj1)-m+1):
                          ##Temporary Result
                          TepRes = 0
                          if UseInp == 1 or UseInp == 0:
                              TepRes = abs(Subj1.count(Subj1[n:n+m])-Subj2.count(Subj1[n:n+m]))
                              TepRes = (1/(TepRes + 1))/FitSli
                          elif  UseInp == 2:
                              TepRes = abs(Subj1.find(Subj1[n:n+m])-Subj2.find(Subj1[n:n+m]))
                              TepRes = (1/(TepRes + 1))/FitSli
                          else:
                              TepRes = abs(Subj1.count(Subj1[n:n+m])-Subj2.count(Subj1[n:n+m]))
                              TepRes = (Ratio*(1/(TepRes+1)))/FitSli
                              TepRes1 = abs(Subj1.find(Subj1[n:n+m])-Subj2.find(Subj1[n:n+m]))
                              TepRes1 = ((1-Ratio)*(1/(TepRes1+1)))/FitSli
                              TepRes = (TepRes + TepRes1)
                          LisRes.append([TepRes, i+2, j, k+2, l])
                      if UseInp == 0:
                        break
                  TextMiniResu.append(LisRes) 
                  TotaSimi()
                  TextMiniResu = []

##The Most Similar Subject                  
def TheMosSimSub(Colum1, Colum2):
  global TextMiniResu, TotSimLis
  Findit = 0
  Colum1, Colum2 = Prepr(Colum1, Colum2)
  for Text1 in Colum1:
    for Subj1 in Text1:
      if Subj1 in Colum2:
        Colum2.remove(Subj1)
        TextMiniResu.append(1)
        continue
      Findit = 0
      for Text2 in Colum2:
        for Subj2 in Text2:
          if Subj1 in Subj2:
            Colum2[Colum2.index(Text2)].remove(Subj2)
            TextMiniResu.append(1)
            Findit = 1
            break
        if Findit == 1:
          break
      if Findit == 0:
        TextMiniResu.append(0)
  TotSimLis = TextMiniResu
  TextMiniResu = []

##The Similarity of Subjects in one row
def OneByOne(Colum1, Colum2, Switch):
  global TotSimLis
  Colum1, Colum2 = Prepr(Colum1, Colum2)
  for Text1 in Colum1:
    Findit = 0
    for Subj1 in Text1:
      if Subj1 == 'nan' or Colum2[Colum1.index(Text1)] == 'nan':
        break
      if Subj1 in str(Colum2[Colum1.index(Text1)]):
        #print(f"\n{Subj1}")
        Findit += 1
    if Findit != 0:
      #Let use len of text or len of List
      LenText = 0
      if Switch == 0:
        LenText = len(Colum2[Colum1.index(Text1)])
      TotSimLis.append(Findit/max(len(Text1),LenText))

def MainViwe(Colum1 = [], Colum2 = [], Switch = 0):
  global TotSimLis
  print("\nYou can choose one of the operations:",
      "\nSimilarity review based on the number of single characters [0]",
      "\nSimilarity review based on the number of same characters   [1]",
      "\nSimilarity review based on the location of same characters [2]",
      "\nSimilarity review based on [1], [2]                        [3]",
      "\nSimilarity review based on same word of all data           [4]",
      "\nSimilarity review based on same word                       [5]",
      "\nEnd the Program                                            [6]\n")
  UseInp = GetInp()
  
  ##Start of Timer 
  ST = time.perf_counter()
  if UseInp != 6:
    if UseInp != 4 and UseInp != 5:
      NumLocSamChr(Colum1, Colum2, 0.8, UseInp)
    elif UseInp == 4:
      TheMosSimSub(Colum1, Colum2)
    elif UseInp == 5:
      OneByOne(Colum1, Colum2, Switch)
    NumSam = len(TotSimLis)
    if NumSam == 0:
       NumSam = 1
    TotalSimi = sum(TotSimLis)/NumSam
    print("\nTotal similarity is: {:.2f}% \n\nRuntime: {:.2f} Sec"\
          .format((TotalSimi*100), (time.perf_counter()-ST)))
    TotSimLis = []
    return TotalSimi
    
TextMiniResu = []
TotSimLis = []

Columns = pandas.read_excel("2008-2021-scopus.xlsx")

#KeywPlu = list(Columns['Index Keywords'])
AuthKey = list(Columns['Author Keywords'])
ArtiTit = list(Columns['Title'])
Abstrac = list(Columns['Abstract'])
#KeywPlu = numpy.reshape(KeywPlu, (len(KeywPlu),1))
AuthKey = numpy.reshape(AuthKey, (len(AuthKey),1))
ArtiTit = numpy.reshape(ArtiTit, (len(ArtiTit),1))
Abstrac = numpy.reshape(Abstrac, (len(Abstrac),1))

SumOfResu = 0

#SumOfResu += MainViwe(KeywPlu, AuthKey, 0)
SumOfResu += MainViwe(AuthKey, ArtiTit, 1)
SumOfResu += MainViwe(AuthKey, Abstrac, 1)

print("Totle similarity of 2 approaches is: {:.2f}".format((SumOfResu*100)/2))
