import requests, openpyxl
from bs4 import BeautifulSoup


def runscrapper():
    try:
        #getting access to the web page
        source = requests.get('https://www.imdb.com/chart/top')
        print(f'Status code = {source.raise_for_status()}')

        #parcing the html page and accessing element by element
        soup = BeautifulSoup(source.text, 'html.parser')
        movies = soup.find('tbody', class_='lister-list').find_all('tr')
        for movie in movies:
            name = movie.find('td', class_='titleColumn').a.text
            rank = int(movie.find('td', class_='titleColumn').text.strip().split('.')[0])
            year = int(movie.find('td', class_='titleColumn').span.text.strip('()'))
            rating = float(movie.find('td', class_='ratingColumn imdbRating').text.strip())
            sheet.append([rank, name, year, rating])
            excel.save('IMDB_Movie_Ratings.xlsx')
    except Exception as e:
        print(e)

if __name__ == "__main__":

    #creating a new excel file and renaming the sheet to the desired name
    excel = openpyxl.Workbook()
    sheet = excel.active
    sheet.Title = 'Top Rated IMDB Movies'
    sheet.append(['Rank', 'Movie Name',  'Year of Release', 'Rating'])
    runscrapper()
