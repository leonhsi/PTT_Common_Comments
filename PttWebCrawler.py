import re
import urllib
import requests
import openpyxl
from multiprocessing import Pool
from requests_html import HTML
from utils import pretty_print  # noqa
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook 


def fetch(url):
    ''' Step-1: send a request and fetch the web page.
    '''
    #response = requests.get(url, cookies={'over18': 'yes'})
    response2 = requests.get(url, cookies={'over18': 'yes'}, params={'q': '疫情 recommend:30'})
    return response2


def parse_article_entries(doc):
    ''' Step-2: parse the post entries on the source string.
    '''
    html = HTML(html=doc)
    post_entries = html.find('div.r-ent')
    return post_entries


def parse_article_push_content(url):
    push_content = []  #list

    resp = requests.get(url, cookies={'over18':'yes'})
    soup = BeautifulSoup(resp.text,'lxml')
    articles = soup.find_all('div','push')
    #print("articles: ",articles)

    for article in articles:
        if type(article.find('span','f3 push-content'))!=type(None):
            data = article.find('span', 'f3 push-content').getText().replace(':','').strip()
            #print(data)
            push_content.append(data)

    return push_content

def parse_article_meta(ent):
    ''' Step-3: parse the metadata in article entry
    '''
    meta = {
        'title': ent.find('div.title', first=True).text,
        'push': ent.find('div.nrec', first=True).text,
        'date': ent.find('div.date', first=True).text,
    }
    
    try:
        meta['author'] = ent.find('div.author', first=True).text
        meta['link'] = ent.find('div.title > a', first=True).attrs['href']
        
    except AttributeError:
        if '(本文已被刪除)' in meta['title']:
            match_author = re.search('\[(\w*)\]', meta['title'])
            if match_author:
                meta['author'] = match_author.group(1)
        elif re.search('已被\w*刪除', meta['title']):
            match_author = re.search('\<(\w*)\>', meta['title'])
            if match_author:
                meta['author'] = match_author.group(1)
    
    
    #print(meta['link'])
    article_url = urllib.parse.urljoin(domain, meta['link'])
    #print("article_url: ",article_url)
    push_content = parse_article_push_content(article_url)
    #print("push content: ",push_content)

    meta['push_content'] = push_content
    

    
    return meta


def get_metadata_from(url):
    ''' Step-4: parse the link of previous link.
    '''

    def parse_next_link(doc):
        ''' Step-4a: parse the link of previous link.
        '''
        html = HTML(html=doc)
        controls = html.find('.action-bar a.btn.wide')
        link = controls[1].attrs.get('href')
        return urllib.parse.urljoin(domain, link)

    resp = fetch(url)
    
    post_entries = parse_article_entries(resp.text)
    next_link = parse_next_link(resp.text)
    #print(next_link)

    metadata = [parse_article_meta(entry) for entry in post_entries]
    return metadata, next_link


def get_paged_meta(url, num_pages):
    ''' Step-4-ext: collect pages of metadata starting from url.
    '''
    collected_meta = []

    for _ in range(num_pages):
        posts, link = get_metadata_from(url)
        collected_meta += posts
        url = urllib.parse.urljoin(domain, link)

    return collected_meta


    
domain = 'https://www.ptt.cc/'
start_url = 'https://www.ptt.cc/bbs/Gossiping/search'



if __name__ == '__main__':
    #creat execel workbook
    wb = Workbook() 
    ws = wb.active
    ws.append(["推文數","文章標題","日期","推文"])

    #load data to excel
    metadata = get_paged_meta(start_url, num_pages=28)
    for meta in metadata:
        for i in range(len(meta['push_content'])):
            pretty_print(meta['push'], meta['title'], meta['date'],meta['push_content'][i])
            ws.append([meta['push'],meta['title'],meta['date'],meta['push_content'][i]])
    
    wb.save('PttWebCrawler_data_num28.xlsx') 
   