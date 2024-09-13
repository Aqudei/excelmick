from bs4 import BeautifulSoup

with open("./result.html",'rt') as infile:
    html_content = infile.read()

# Parse the HTML using BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')
element = soup.select_one(".search-results")

# Extract name
name = element.find('h4').get_text(strip=False).split('<br>')[0]

# Extract phone number
phone = element.find('td', string='Phone ').find_next_sibling('td').get_text()

# Extract email
email = element.find('td', string='Email ').find_next_sibling('td').get_text()

# Extract address
address = element.find('td', string='Address').find_next_sibling('td').get_text(separator=', ')

# Extract types
types = [span.get_text() for span in element.find_all('div', class_='types')[0].find_all('span')]

# Output the parsed data
import re

name = re.sub(r"\s+"," ",name)
print(f"Name: {name}")
print(f"Phone: {phone}")
print(f"Email: {email}")
print(f"Address: {address}")
print(f"Types: {', '.join(types)}")