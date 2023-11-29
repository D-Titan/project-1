import kivy
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout

import requests
from bs4 import BeautifulSoup
import pandas as pd
import os



def extract_data(stock_symbol, c):
    url = f"https://www.screener.in/company/{stock_symbol}"
    if c.lower() == 'c':
        url += '/consolidated/'
    elif c.lower()=='s':
      url += '/'
    else:
      url += '/'

    response = requests.get(url)
    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    def html_table(Id,Soup):

      #Finding Sections
      section = Soup.find_all(id=Id)
      soup2 = BeautifulSoup(str(section),'html.parser')

      #Table Heading in a list
      thead = soup2.find_all("thead")
      souph = BeautifulSoup(str(thead),'html.parser')
      l_thead = souph.find_all("th")

      #Table body in a list that contains list of every rows
      tbody = soup2.find_all("tbody")

      soupb = BeautifulSoup(str(tbody),'html.parser')
      l_tbody_tr = soupb.find_all("tr")

      tr = []
      for i in l_tbody_tr:
        soup_tr = BeautifulSoup(str(i),'html.parser')
        tr.append(soup_tr.find_all("td"))


      #if thead is not available
      if thead==[]:
        l_thead = soup2.find_all("th")

      s=0
      for i in l_thead:
        l_thead[s]= i.text.replace("\n",'').replace(" ",'').replace("?",'').replace("Cr.",'').replace(",",'')
        s+=1
      m=0
      for i in tr:
        for j in range(len(i)):
            tr[m][j]=i[j].text.replace("\n",'').replace(" ",'').replace("?",'').replace("Cr.",'').replace(",",'')
        m+=1



      return pd.DataFrame(tr,columns=l_thead)

    """ FOR Basic RATIOS TABLE """

    # Find all list tags.
    lt = soup.find_all(id="top-ratios")
    soup1 = BeautifulSoup(str(lt), "html.parser")

    #soup = BeautifulSoup(html_content, "html.parser")
    lnm= soup1.find_all(class_="name")
    lvl=soup1.find_all(class_="nowrap value")

    dl = {} # Create a list to store the list values.

    # Iterate through the list tags and get the text.
    s=0
    for i in lnm:
        dl[i.text.replace("\n",'').replace(" ",'')] = lvl[s].text.replace("\n",'').replace(" ",'').replace("?",'').replace("Cr.",'').replace(",",'')
        s+=1

    df1=pd.DataFrame(dl,index=["Values"]) #BASIC RATIOS TABLE

    # Extract data from the website using BeautifulSoup and Pandas
    df2 = html_table('profit-loss', soup)  # PROFIT & LOSS TABLE
    df3 = html_table('balance-sheet', soup)  # BALANCE SHEET TABLE
    dlf4=pd.read_html('https://www.indiainfoline.com/markets/sector-valuation')[0]

    # Combine the data into a single DataFrame
    data = [df1, df2, df3,dlf4]

    return data


def save_to_excel(data,stock_symbol):
    # Create an Excel workbook and worksheets
    download_dir = os.path.expanduser('/storage/emulated/0/Download')

    # Create the Excel file
    file_path = os.path.join(download_dir, stock_symbol+"-final.xlsx")
    writer = pd.ExcelWriter(file_path, engine="xlsxwriter")
    data[0].to_excel(writer, sheet_name='Sheet1',index=False)
    data[1].to_excel(writer, sheet_name='Sheet1',startrow=len(data[0])+3,index=False)
    data[2].to_excel(writer, sheet_name='Sheet1',startrow=len(data[0])+len(data[1])+6,index=False)
    data[3].to_excel(writer, sheet_name='Sheet1',startrow=len(data[0])+len(data[1])+len(data[2])+9, index=False)
    writer.close()


class MyApp(App):
    def build(self):
        self.stock_symbol = TextInput(multiline=False)
        self.c = TextInput(multiline=False)
        self.button = Button(text="Extract and Save")
        self.button.bind(on_press=self.extract_and_save)

        layout = BoxLayout(orientation='vertical')
        layout.add_widget(Label(text="Stock Symbol:"))
        layout.add_widget(self.stock_symbol)
        layout.add_widget(Label(text="Consolidated or Standalone (c/s):"))
        layout.add_widget(self.c)
        layout.add_widget(self.button)

        return layout

    def extract_and_save(self, instance):
        stock_symbol = self.stock_symbol.text.upper()
        c = self.c.text

        # Extract data from the website and save it to an Excel spreadsheet
        data = extract_data(stock_symbol, c)
        save_to_excel(data,stock_symbol)

        # Clear the text inputs for next use
        self.stock_symbol.text = ""
        self.c.text = ""
        

if __name__ == "__main__":
    MyApp().run()
