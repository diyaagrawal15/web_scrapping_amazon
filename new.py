import openpyxl
from bs4 import BeautifulSoup

# Read the Amazon.html file
with open("Amazon.html", "r", encoding="utf-8") as file:
    html_content = file.read()

# Function to parse the HTML and extract required information
def parse_amazon_html(html):
    soup = BeautifulSoup(html, 'html.parser')
    product_divs = soup.find_all('div', class_='a-section a-spacing-small a-spacing-top-small')

    products_data = []
    for div in product_divs:
        product_name = div.find('h2', class_='a-size-mini a-spacing-none a-color-base s-line-clamp-2')
        product_name = product_name.text.strip() if product_name else " "

        product_price = div.find('span', class_='a-price-whole')
        product_price = product_price.text.strip() if product_price else " "

        product_review = div.find('span', class_='a-icon-alt')
        product_review = product_review.text.strip() if product_review else " "

        no_of_review = div.find('span', class_='a-size-base s-underline-text')
        no_of_review = no_of_review.text.strip() if no_of_review else " "

        product_url = div.find('a', class_='a-link-normal s-underline-text s-underline-link-text s-link-style')
        product_url = product_url['href'] if product_url else " "

        products_data.append((product_name, product_price, product_review, no_of_review, product_url))

    return products_data

# Function to write the extracted data to an Excel file
def write_to_excel(products_data):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    headers = ["Product Name", "Product Price", "Product Review", "Number of Reviews", "Product URL"]
    sheet.append(headers)

    for product in products_data:
        sheet.append(product)

    workbook.save("amazon_products.xlsx")

# Main function
if __name__ == "__main__":
    parsed_data = parse_amazon_html(html_content)
    write_to_excel(parsed_data)
    print("Data written to amazon_products.xlsx")
