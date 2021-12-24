import os
import sys
import time
import openpyxl
import operator
import matplotlib.pyplot as plt
from sklearn import feature_extraction
from sklearn.feature_extraction.text import TfidfTransformer
from sklearn.feature_extraction.text import CountVectorizer
from ckiptagger import data_utils, construct_dictionary, WS, POS, NER
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'


#get data from excel file
def get_data_iter(filename,sheet,max_row,min_row,max_col,min_col):
    wb = openpyxl.load_workbook(filename)
    sheet = wb[sheet]
    cell = sheet.iter_rows(max_row=max_row, min_row=min_row, max_col=max_col, min_col=min_col)
    all_rows=[]
    str=""

    for row in cell:
        rows = []
        for c in row:
            if type(c.value)== type(str) :    #in case of NoneType
                str = str + c.value
            
    all_rows.append(str)
    wb.close()
    return all_rows

#slice words by Ckiptagger
def get_WS(length):

    ws = WS("./data")
    pos = POS("./data")
    ner = NER("./data")

    rows = get_data_iter('PttWebCrawler_data_0227_0413.xlsx','Sheet',max_row=length,min_row=2,max_col=4,min_col=4)
    sentence_list = rows

    word_sentence_list = ws(sentence_list)
    pos_sentence_list = pos(word_sentence_list)
    entity_sentence_list = ner(word_sentence_list, pos_sentence_list)

    #release memory
    del ws
    del pos
    del ner

    #print result of Ckiptagger:
    def print_word_pos_sentence(word_sentence, pos_sentence):
        assert len(word_sentence) == len(pos_sentence)
        for word, pos in zip(word_sentence, pos_sentence):
            print(f"{word}({pos})", end="\u3000")
        print()
        return
    
    for i, sentence in enumerate(sentence_list):
        print()
        #print(f"'{sentence}'")
        print_word_pos_sentence(word_sentence_list[i],  pos_sentence_list[i])
        for entity in sorted(entity_sentence_list[i]):
            print(entity)


    return word_sentence_list[0]

#evaluate frequency using tf-idf
def get_tfidf(length, words_number):

    text = ' '.join(get_WS(length))
    corpus = []
    corpus.append(text)

    vectorizer = CountVectorizer()
    transformer = TfidfTransformer()
    tfidf = transformer.fit_transform(vectorizer.fit_transform(corpus))
    word = vectorizer.get_feature_names()
    weight = tfidf.toarray()

    ##print tfidf result:
    #for i in range(len(weight)):
    #    print (u"---------",i,u"------------")
    #    for j in range(len(word)):
    #        print (word[j],weight[i][j])

    
    
    #sort tfidf result
    pair_list = [ ((word[i]), (weight[0][i]))  for i in range(len(word)) ] 
    sorted_pair_list = sorted(pair_list, key=operator.itemgetter(1), reverse=True)

    seg_list = []       
    for i in range(words_number):
        seg_list.append(sorted_pair_list[i])

    for i in range(words_number):
        print(seg_list[i])
    

    #creat dictionary to store segment list as {'words': 'frequency'}
    d={}
    for x, y in seg_list:
        d[x] = y

    return d 
    

if __name__ == "__main__":
    tStart = time.time()

    data_length =  106       #the number of columns(pushes) from excel
    words_number = 200        #the number of words showing on the wordcloud

    #creat word cloud
    d = get_tfidf(data_length,words_number)

    font = r'msjh.ttc'
    wordcloud = WordCloud(background_color="black",font_path=font,collocations=False, width=2400, height=2400, margin=2)
    wordcloud.generate_from_frequencies(d)
    plt.figure( figsize=(20,10), facecolor='k')
    plt.imshow(wordcloud,interpolation='bilinear')
    plt.axis("off")
    plt.tight_layout(pad=0)
    plt.show()      #show pics


    tEnd = time.time()
    print("******total time cost****** ", tEnd-tStart, " s")


    
