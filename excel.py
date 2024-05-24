import openpyxl
import pandas as pd
from openai import OpenAI


client = OpenAI()

df = pd.read_excel('ExtraData.xlsx')

wb = openpyxl.Workbook()
ws = wb['Sheet']
ws.append(['reference_id', 'title', 'author', 'published_year', 'author_year_of_death'])

try:
    for i, record in enumerate(df.values):
        print(i)
        _id, title, author, published_year, author_year_of_death = record
        #
        if i >= 0:
            published_year_query = f'Return only the year "{title}" by {author} was published, or "XXXX" if unknown.'
            published_year_completion = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {
                        "role": "system",
                        "content": published_year_query
                    },
                ]
            )
            published_year = published_year_completion.choices[0].message.content
            #
            if author not in ['Anonymous', 'Various', '#N/A', '']:
                author_year_of_death_query = f'Provide only the year of death for {author}, author of {title}. If the author is still alive, return "2025", if you do not know it return "YYYY".'
                author_year_of_death_completion = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {
                            "role": "system",
                            "content": author_year_of_death_query
                        },
                    ]
                )
                author_year_of_death = author_year_of_death_completion.choices[0].message.content
            else:
                author_year_of_death = '----'
        ws.append([_id, title, author, published_year, author_year_of_death])
except Exception as e:
    print(str(e))

wb.save('Guttenberg-22.05.xlsx')
