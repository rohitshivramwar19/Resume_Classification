#!/usr/bin/env python
# coding: utf-8

# In[9]:


#pip install textract
#pip install docx
#pip install python-docx 
#pip install antiword
#pip install pdftotext
#pip install PyPDF2


# In[16]:



import nltk
import pandas as pd
import PyPDF2
import matplotlib.pyplot as plt
import re
from nltk.stem import WordNetLemmatizer
from nltk.corpus import stopwords, wordnet
from wordcloud import WordCloud, STOPWORDS
import seaborn as sns
import numpy as np
import csv
import warnings
import os
import docx2txt
import textract
import antiword
from docx import Document
import win32com.client as win32
from sklearn.ensemble import AdaBoostClassifier, GradientBoostingClassifier
from sklearn.linear_model import LogisticRegression
import seaborn as sns
from wordcloud import WordCloud

from sklearn.preprocessing import LabelEncoder


warnings.filterwarnings('ignore')


# In[17]:


os.listdir(r'C:\Users\ROHIT\Desktop\PROJECT 2 (FINAL)\Resumes-20211103T133301Z-001.zip (Unzipped Files)\Resumes')


# In[18]:


def convert_doc_to_docx(doc_file):
    # Initialize Word application
    word = win32.Dispatch('Word.Application')

    # Set visibility to False so that the Word application is not shown
    word.Visible = False

    # Open the .doc file
    doc = word.Documents.Open(doc_file)

    # Get the file path and name without the extension
    file_path, file_name = os.path.split(doc_file)
    file_name_without_ext = os.path.splitext(file_name)[0]

    # Create the .docx file path
    docx_file = os.path.join(file_path, f"{file_name_without_ext}.docx")

    # Save the .doc file as .docx format
    doc.SaveAs2(docx_file, FileFormat=16)  # FileFormat=16 specifies .docx format

    # Close the .doc file
    doc.Close()

    # Quit Word application
    word.Quit()

    print(f"File converted: {docx_file}")


# In[20]:


directory = r'C:\Users\ROHIT\Desktop\PROJECT 2 (FINAL)\Resumes-20211103T133301Z-001.zip (Unzipped Files)\Resumes'
file_path = []
category = []

for filename in os.listdir(directory):
    if filename.endswith('.docx'):
        path = os.path.join(directory, filename)
        text = textract.process(path).decode('utf-8')
        file_path.append(text)
        category.append('React JS Developer Resume')

    elif filename.endswith('.pdf'):
        path = os.path.join(directory, filename)
        pdf_file = open(path, 'rb')
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
        file_path.append(text)
        category.append('React JS Developer Resume')

        pdf_file.close()
        
    elif filename.endswith('.doc'):
        path = os.path.join(directory, filename)
        convert_doc_to_docx(path)
        docx_path = os.path.join(directory, f"{os.path.splitext(filename)[0]}.docx")
        text = textract.process(docx_path).decode('utf-8')
        file_path.append(text)
        category.append('React JS Developer Resume')


# In[21]:


directory = r'C:\Users\ROHIT\Desktop\PROJECT 2 (FINAL)\Resumes-20211103T133301Z-001.zip (Unzipped Files)\Resumes\Peoplesoft resumes'
file_path1 = []
category1 = []

for filename in os.listdir(directory):
    if filename.endswith('.docx'):
        path = os.path.join(directory, filename)
        text = textract.process(path).decode('utf-8')
        file_path1.append(text)
        category1.append('PeopleSoft Resume')

    elif filename.endswith('.pdf'):
        path = os.path.join(directory, filename)
        pdf_file = open(path, 'rb')
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
        file_path1.append(text)
        category1.append('PeopleSoft Resume')

        pdf_file.close()
        
    elif filename.endswith('.doc'):
        path = os.path.join(directory, filename)
        convert_doc_to_docx(path)
        docx_path = os.path.join(directory, f"{os.path.splitext(filename)[0]}.docx")
        text = textract.process(docx_path).decode('utf-8')
        file_path.append(text)
        category.append('PeopleSoft Resume')


# In[22]:


file_path2 = []
category2 = []
directory = r'C:\Users\ROHIT\Desktop\PROJECT 2 (FINAL)\Resumes-20211103T133301Z-001.zip (Unzipped Files)\Resumes\workday resumes'

for filename in os.listdir(directory):
    if filename.endswith('.docx'):
        path = os.path.join(directory, filename)
        text = textract.process(path).decode('utf-8')
        file_path2.append(text)
        category2.append('Workday Resume')

    elif filename.endswith('.pdf'):
        path = os.path.join(directory, filename)
        pdf_file = open(path, 'rb')
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
        file_path2.append(text)
        category2.append('Workday Resume')

        pdf_file.close()
        
    elif filename.endswith('.doc'):
        path = os.path.join(directory, filename)
        convert_doc_to_docx(path)
        docx_path = os.path.join(directory, f"{os.path.splitext(filename)[0]}.docx")
        text = textract.process(docx_path).decode('utf-8')
        file_path.append(text)
        category.append('Workday Resume')


# In[23]:


file_path3 = []
category3 = []
directory = r'C:\Users\ROHIT\Desktop\PROJECT 2 (FINAL)\Resumes-20211103T133301Z-001.zip (Unzipped Files)\Resumes\SQL Developer Lightning insight'

for filename in os.listdir(directory):
    if filename.endswith('.docx'):
        path = os.path.join(directory, filename)
        text = textract.process(path).decode('utf-8')
        file_path3.append(text)
        category3.append('SQL Developer Lightning Insight Resume')

    elif filename.endswith('.pdf'):
        path = os.path.join(directory, filename)
        pdf_file = open(path, 'rb')
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
        file_path3.append(text)
        category3.append('SQL Developer Lightning Insight Resume')

        pdf_file.close()
        
    elif filename.endswith('.doc'):
        path = os.path.join(directory, filename)
        convert_doc_to_docx(path)
        docx_path = os.path.join(directory, f"{os.path.splitext(filename)[0]}.docx")
        text = textract.process(docx_path).decode('utf-8')
        file_path.append(text)
        category.append('SQL Developer Lightning Insight Resume')


# In[24]:


data_1 = pd.DataFrame(data = file_path , columns = ['Resumes'])
data_1['category'] = category
data_1


# In[25]:


data_2 = pd.DataFrame(data = file_path1 , columns = ['Resumes'])
data_2['category_1'] = category1
data_2


# In[26]:


data_3 = pd.DataFrame(data = file_path2 , columns = ['Resumes'])
data_3['category_2'] = category2
data_3


# In[27]:


data_4 = pd.DataFrame(data = file_path3 , columns = ['Resumes'])
data_4['category_3'] = category3
data_4


# In[28]:


df = data_1.append([data_2, data_3, data_4], ignore_index = True)
df['Category'] = category + category1 + category2 + category3
df


# In[29]:


df.drop(columns=['category','category_1','category_2','category_3'], inplace=True)


# In[30]:


df


# In[31]:


df.to_csv('converted_resumes.csv', index=False)


# In[ ]:





# In[32]:


df = pd.read_csv('converted_resumes.csv')


# In[33]:


df.tail(10)


# In[34]:


df.describe()


# In[35]:


df.info()


# In[36]:


df[df.duplicated("Resumes")]


# In[37]:


def clean_text(text):
    

    # Remove links
    text = re.sub(r'http\S+', ' ', text)
    
    # Remove punctuations
    text = re.sub(r'[^\w\s]', ' ', text)
    
    # Remove non-english alphabets
    text = ''.join([i for i in text if i.isalpha() or i.isspace()])
    
    # Remove numbers
    text = re.sub(r'\d+', ' ', text)
    
    text = text.lower()

    # Tokenize the text
    tokens = nltk.word_tokenize(text)

    # Remove stopwords and lemmatize the remaining words
    lemma= WordNetLemmatizer()
    stop_words = nltk.corpus.stopwords.words("english")


    nouns = [token for token in tokens if nltk.pos_tag([token])[0][1] == "NOUN"]
    tokens = [lemma.lemmatize(i) for i in tokens if not i in stop_words and i not in nouns]
    
    # Join the tokens back into a string
    text = ' '.join(tokens)

    return text


# In[38]:


# applying function on resumes

df['Resumes']= df['Resumes'].apply(clean_text)
df


# In[39]:


df['Category'].value_counts()


# In[40]:


plt.rcParams['figure.figsize']= (12,6)
sns.set_style(style='darkgrid')


# In[41]:


df['Category'].value_counts().plot(kind='pie', autopct='%0.2f%%') 
plt.show()


# - Here the data is approximately balanced, so we may not face multicoliniarity issue.

# In[42]:


workday = df[df['Category']=='Workday Resume']
reactjs = df[df['Category']=='React JS Developer Resume']
sqldeveloper = df[df['Category']=='SQL Developer Lightning Insight Resume']
peoplesoft = df[df['Category']=='PeopleSoft Resume']


# In[43]:


df['Length']= df['Resumes'].apply(lambda x:len(nltk.word_tokenize(x)))


# In[44]:


sns.displot(df['Length'])
plt.show()


# - this way of analysis is not showing any proper result.
# - Most of the resumes are consist of 200 and 400

# In[45]:


from nltk.corpus import stopwords
import string


# In[46]:


oneSetOfStopWords = set(stopwords.words('english')+['``',"''"])
totalWords =[]
Sentences = df['Resumes'].values
cleanedSentences = ""
for i in range(0,77):
    cleanedText = clean_text(Sentences[i])
    cleanedSentences += cleanedText
    requiredWords = nltk.word_tokenize(cleanedText)
    for word in requiredWords:
        if word not in oneSetOfStopWords and word not in string.punctuation:
            totalWords.append(word)
    
wordfreqdist = nltk.FreqDist(totalWords)
mostcommon = wordfreqdist.most_common(200)
print(mostcommon)


# In[47]:


words0= []

for word, count in mostcommon:
    words0.append(word)


# In[48]:


WORDCLOUD_COLOR_MAP = 'tab10_r'


# In[49]:


def wordcl(data, title):
    stop = STOPWORDS
    wc = WordCloud(height=2000, width= 4000, colormap=WORDCLOUD_COLOR_MAP, stopwords=stop).generate(data)
    plt.imshow(wc)
    plt.axis('off')
    plt.title(title)


# In[50]:


wordcl(" ".join(words0), "Wordcloud")


# In[ ]:





# In[ ]:





# In[51]:


oneSetOfStopWords = set(stopwords.words('english')+['``',"''"])
totalWords =[]
Sentences = workday['Resumes'].values
cleanedSentences = ""
for i in range(len(workday)):
    cleanedText = clean_text(Sentences[i])
    cleanedSentences += cleanedText
    requiredWords = nltk.word_tokenize(cleanedText)
    for word in requiredWords:
        if word not in oneSetOfStopWords and word not in string.punctuation:
            totalWords.append(word)
    
wordfreqdist = nltk.FreqDist(totalWords)
mostcommon = wordfreqdist.most_common(50)
print(mostcommon)


# In[ ]:





# In[ ]:





# In[52]:


oneSetOfStopWords = set(stopwords.words('english')+['``',"''"])
totalWords =[]
Sentences = reactjs['Resumes'].values
cleanedSentences = ""
for i in range(len(reactjs)):
    cleanedText = clean_text(Sentences[i])
    cleanedSentences += cleanedText
    requiredWords = nltk.word_tokenize(cleanedText)
    for word in requiredWords:
        if word not in oneSetOfStopWords and word not in string.punctuation:
            totalWords.append(word)
    
wordfreqdist = nltk.FreqDist(totalWords)
mostcommon = wordfreqdist.most_common(50)
print(mostcommon)


# In[ ]:





# In[ ]:





# In[53]:


oneSetOfStopWords = set(stopwords.words('english')+['``',"''"])
totalWords =[]
Sentences = sqldeveloper['Resumes'].values
cleanedSentences = ""
for i in range(len(sqldeveloper)):
    cleanedText = clean_text(Sentences[i])
    cleanedSentences += cleanedText
    requiredWords = nltk.word_tokenize(cleanedText)
    for word in requiredWords:
        if word not in oneSetOfStopWords and word not in string.punctuation:
            totalWords.append(word)
    
wordfreqdist = nltk.FreqDist(totalWords)
mostcommon = wordfreqdist.most_common(50)
print(mostcommon)


# In[ ]:





# In[54]:


oneSetOfStopWords = set(stopwords.words('english')+['``',"''"])
totalWords =[]
Sentences = peoplesoft['Resumes'].values
cleanedSentences = ""
for i in range(len(peoplesoft)):
    cleanedText = clean_text(Sentences[i])
    cleanedSentences += cleanedText
    requiredWords = nltk.word_tokenize(cleanedText)
    for word in requiredWords:
        if word not in oneSetOfStopWords and word not in string.punctuation:
            totalWords.append(word)
    
wordfreqdist = nltk.FreqDist(totalWords)
mostcommon = wordfreqdist.most_common(50)
print(mostcommon)


# In[55]:


categories = np.sort(df['Category'].unique())
categories


# In[56]:


df_categories = [df[df['Category'] == category].loc[:, ['Resumes', 'Category']] for category in categories]


# In[57]:


def wordcloud(df):
    txt = ' '.join(txt for txt in df['Resumes'])
    wordcloud = WordCloud(
        height=2000,
        width=4000,
        colormap=WORDCLOUD_COLOR_MAP
    ).generate(txt)

    return wordcloud


# In[58]:


WORDCLOUD_COLOR_MAP = 'tab10_r'
plt.figure(figsize=(64, 56))

for i, category in enumerate(categories):
    wc = wordcloud(df_categories[i])

    plt.subplot(4, 1, i + 1).set_title(category,fontsize=30,fontweight= 'bold')
    plt.imshow(wc)
    plt.axis('off')
    plt.plot()

plt.show()
plt.close()


# # Unique Words

# In[59]:


unique = df.groupby('Category')['Resumes'].apply(lambda x: ' '.join(x))

# Generate a word cloud for each group
for category, text in unique.items():
    wordcloud = WordCloud(height=1500,width=2000,colormap=WORDCLOUD_COLOR_MAP).generate(text)

    WORDCLOUD_COLOR_MAP = 'tab10_r'
    plt.figure(figsize=(15, 25))

    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis("off")
    plt.title(category,fontsize=28,fontweight= 'bold')
    plt.show()


# In[60]:


from textblob import TextBlob


# In[ ]:


#TextBlob(df['Resumes'][77]).ngrams(1)[:20]


# In[ ]:


#TextBlob(df['Resumes'][77]).ngrams(2)[:50]


# In[61]:


Stopwords = (set(nltk.corpus.stopwords.words("english")))


# In[ ]:





# In[62]:


def preprocess(sentence):
    sentence = str(sentence)
    sentence = sentence.lower()
    sentence = sentence.replace('using', ' ')
    sentence = sentence.replace('work', ' ')
    sentence = sentence.replace('used', ' ')
    sentence = sentence.replace('good', ' ')
    sentence = sentence.replace('various', ' ')
    sentence = sentence.replace('project', ' ')
    sentence = sentence.replace('experience', ' ')
    sentence = sentence.replace('application', ' ')
    sentence = sentence.replace('involved', ' ')
    
    # Tokenize the text
    tokens = nltk.word_tokenize(sentence)

    # Remove stopwords and lemmatize the remaining words
    lemma= WordNetLemmatizer()
    stop_words = nltk.corpus.stopwords.words("english")


    nouns = [token for token in tokens if nltk.pos_tag([token])[0][1] == "NOUN"]
    tokens = [lemma.lemmatize(i) for i in tokens if not i in stop_words and i not in nouns]
    
    # Join the tokens back into a string
    text = ' '.join(tokens)

    return text


# In[63]:


df['Resumes'] = df['Resumes'].apply(preprocess)


# In[64]:


target_words=['used','responsibility', 'responsible', 'university','various','involved',
              'etc', 'school', 'college', 'engineering','profile', 'worked', 'target',
              'system','report', 'knowledge']


# In[65]:


def replace_target_words(target_words, sentences):
    for sentence in sentences:
        
        for target_word in target_words:
            sentence = re.sub(target_word, ' ', sentence)

    return sentences


# In[66]:


df['Resumes'] = replace_target_words(target_words, df['Resumes'])


# In[ ]:


#num=int(input("there are only 4 categories and index start with 0 so give number between 0 to 3 :"))
#df_categories[num]


# In[67]:


df_categories = [df[df['Category'] == category].loc[:, ['Resumes', 'Category']] for category in categories]


# In[68]:


def wordcloud(df):
    txt = ' '.join(txt for txt in df['Resumes'])
    wordcloud = WordCloud(
        height=2000,
        width=4000,
        colormap=WORDCLOUD_COLOR_MAP
    ).generate(txt)

    return wordcloud


# In[69]:


WORDCLOUD_COLOR_MAP = 'tab10_r'
plt.figure(figsize=(64, 56))

for i, category in enumerate(categories):
    wc = wordcloud(df_categories[i])

    plt.subplot(4, 1, i + 1).set_title(category,fontsize=30,fontweight= 'bold')
    plt.imshow(wc)
    plt.axis('off')
    plt.plot()

plt.show()
plt.close()


# # Model Building

# In[70]:


from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import GaussianNB
from sklearn.metrics import classification_report, confusion_matrix, accuracy_score, f1_score, precision_score, recall_score
from sklearn.svm import SVC


# In[71]:


def predict(model):
    model = model.fit(xtrain, ytrain)
    ypred = model.predict(xtest)
    
    trainac = model.score(xtrain, ytrain)
    testac = model.score(xtest, ytest)
    
    print(f"Train accuracy {trainac}\nTest accuracy {testac}")
    
    print(classification_report(ytest, ypred))


# In[72]:


from sklearn.ensemble import AdaBoostClassifier, GradientBoostingClassifier, RandomForestClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier


# In[73]:


tf =  TfidfVectorizer(ngram_range=(2,2))
x = tf.fit_transform(df['Resumes'])

x = pd.DataFrame(x.toarray(), columns=tf.get_feature_names_out())
x


# In[74]:


df['Category'].unique()


# In[75]:


df.to_csv("clean_resume.csv")


# In[76]:


label= LabelEncoder()

df['labels'] = label.fit_transform(df['Category'])


# In[77]:


df['labels'].unique()


# In[78]:


df


# In[79]:


y = df['labels']


# In[80]:


xtrain,xtest,ytrain,ytest = train_test_split(x,y,test_size=0.2,random_state=1)


# In[81]:


predict(LogisticRegression())


# In[82]:


predict(DecisionTreeClassifier())


# In[83]:


predict(RandomForestClassifier())


# In[84]:


predict(AdaBoostClassifier())


# In[85]:


predict(GradientBoostingClassifier())


# In[86]:


predict(GaussianNB())


# In[87]:


predict(SVC())


# In[88]:


predict(KNeighborsClassifier())


# # Unigram Model

# In[89]:


tf =  TfidfVectorizer(ngram_range=(1,1))
x = tf.fit_transform(df['Resumes'])

x = pd.DataFrame(x.toarray(), columns=tf.get_feature_names_out())
x


# In[90]:


y = df['labels']


# In[91]:


xtrain,xtest,ytrain,ytest = train_test_split(x,y,test_size=0.2,random_state=1)


# In[92]:


predict(LogisticRegression())


# In[93]:


predict(AdaBoostClassifier())


# In[94]:


predict(GradientBoostingClassifier())


# In[95]:


predict(GaussianNB())


# In[96]:


predict(DecisionTreeClassifier())


# In[97]:


predict(RandomForestClassifier())


# In[ ]:





# # Pipeline

# In[98]:


from sklearn.pipeline import Pipeline


# In[99]:


data= pd.read_csv(r'clean_resume.csv', usecols=['Resumes','Category'])


# In[100]:


data


# In[101]:


xtrain,xtest,ytrain,ytest = train_test_split(data['Resumes'],data['Category'],test_size=0.2,random_state=1)


# In[102]:


# Define pipeline
model = Pipeline([
    ('tfidf', TfidfVectorizer(ngram_range=(2,2))),
    ('model', GradientBoostingClassifier(n_estimators=200, random_state=1))
])


# In[103]:


model.fit(xtrain, ytrain)       # Train model using pipeline
y_pred = model.predict(xtest)   # Evaluate model on testing set
print(classification_report(ytest, y_pred))


# # Model Saving

# In[104]:


import pickle


# In[ ]:


#pickle.dump(model, open('model.pkl','wb'))


# In[ ]:




