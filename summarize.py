# importing libraries 
import docx
from nltk.corpus import stopwords 
from nltk.tokenize import word_tokenize, sent_tokenize 



def getText(filename,keyword):
    doc = docx.Document(filename)
    fullText = []
    keyText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
        if(keyword.lower() in para.text.lower()):
            keyText.append(para.text)
        
    return '\n'.join(keyText),'\n'.join(fullText)

def getKeywordText(keyword,sentences):
    result = []
    for sentence in sentences:
        if keyword.lower() in sentence.lower():
            result.append(sentence)
    return '\n'.join(result)

def getFreqTable(words,stopWords):
    freqTable = dict() 
    for word in words: 
        word = word.lower() 
        if word in stopWords: 
            continue
        if word in freqTable: 
            freqTable[word] += 1
        else: 
            freqTable[word] = 1
    return freqTable

def getSentenceValue(sentences,freqTable):
    sentenceValue = dict() 
    
    for sentence in sentences: 
        for word, freq in freqTable.items(): 
            if word in sentence.lower(): 
                if sentence in sentenceValue: 
                    sentenceValue[sentence] += freq 
                else: 
                    sentenceValue[sentence] = freq 
    return sentenceValue

def getSumValue(sentenceValue):
    sumValues = 0
    for sentence in sentenceValue: 
        sumValues += sentenceValue[sentence]
    return sumValues

def getSummary(sentences,sentenceValue,average):
    summary = '' 
    for sentence in sentences: 
        if (sentence in sentenceValue) and (sentenceValue[sentence] > (1.5 * average)): 
            summary += " " + sentence 
    return summary

def SaveToFile(saveFileName,summary):
    with open(saveFileName,'w',encoding='utf-8') as f:
        f.write(summary)
        f.close()
    print("Successfully saved into " + saveFileName)
    return

filename = input("Please insert content file name : ")
keyword = input("Please insert keywords : ")
keyword_text , fullText = getText(filename,keyword)
text = fullText


# Tokenizing the text 
stopWords = set(stopwords.words("english")) 
words = word_tokenize(text)
sentences = sent_tokenize(text)

words_key = word_tokenize(keyword_text)
sentences_key = sent_tokenize(keyword_text)
# Creating a frequency table to keep the  
# score of each word 
freqTable = getFreqTable(words,stopWords)
freqTable_key = getFreqTable(words_key,stopWords)
# Creating a dictionary to keep the score 
# of each sentence 

sentenceValue = getSentenceValue(sentences,freqTable)

sentenceValue_key = getSentenceValue(sentences_key,freqTable_key)

 
# Average value of a sentence from the original text 
sumValues = getSumValue(sentenceValue)
average = int(sumValues / len(sentenceValue)) 

sumValues_key = getSumValue(sentenceValue_key)
average_key = int(sumValues_key / len(sentenceValue_key))

# Storing sentences into our summary. 
summary = getSummary(sentences,sentenceValue,average)
summary_key = getSummary(sentences_key,sentenceValue_key,average_key)
saveFileName = input('Please insert result summary file name : ')
if(saveFileName.find('.txt') != -1):
    saveFileName = saveFileName.replace('.txt','')

#save to file
SaveToFile(saveFileName + '.txt',summary)
SaveToFile(saveFileName + '_key.txt',summary_key)


