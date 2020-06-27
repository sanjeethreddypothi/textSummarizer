#Source for the text- https://drive.google.com/file/d/1TzBe2kUd7kTpLfh5youSFQPFtFUv1xbU/view
import nltk
import nltk.corpus 
from nltk.corpus import stopwords 
import nltk.tokenize 
from nltk.tokenize import word_tokenize
import pandas as pd
from nltk.tokenize import sent_tokenize
loc = pd.read_excel (r"C:\Users\DELL\Downloads\task.xlsx")
file=pd.DataFrame(loc)
file.drop(file.columns[0:1],axis=1,inplace=True)
#Creating a function to summarize the paragraph
def summarize(i):
#locating each row with the use of .iloc
    text=file.iloc[i,0]
#Following few steps are used to ignore the stopwords used in the paragraph to build a more efficient Frequency Table
    stopWords = set(stopwords.words("english")) 
    tokenz = word_tokenize(text) 
    freqTable = dict() 
    for token in tokenz: 
        token = token.lower() 
        if token in stopWords: 
            continue
        if token in freqTable: 
            freqTable[token] += 1
        else: 
            freqTable[token] = 1   
 #freuency table is created with all the important words used in the paragraph along with respective freuencies 
 #following few steps tokenize each sentence used and mark it using degree of important words used in it to recognize most important sentences
    sentences = sent_tokenize(text)  
    
    sentenceValue = dict() 
   
    for sentence in sentences: 
        for word, freq in freqTable.items(): 
            if word in sentence.lower(): 
                if sentence in sentenceValue: 
                    sentenceValue[sentence] += freq 
                else: 
                    sentenceValue[sentence] = freq 
#Importance values are marked for all the sentences used
#By getting an average score we can reject all the sentecnes with below avergae importance
    sumValues = 0
    for sentence in sentenceValue: 
        sumValues += sentenceValue[sentence]
        avg=sumValues/len(sentenceValue)
#Creating a new summary and adding all the important sentences
    summary = '' 
    for sentence in sentences: 
#making sure only senetences with greater importance are selected
        if (sentence in sentenceValue) and (sentenceValue[sentence] > (1.25* avg)): 
            summary += " " + sentence 
    print(summary) 
    print("length of summary is"+" "+ str(len(summary))+" & length of the total text was" +" "+str(len(text)))
#Using this loop all the cells have been summarized
for i in range(len(file)):
    
    summarize(i)