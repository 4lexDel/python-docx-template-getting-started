from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import jinja2


# our template file containing the word content
template = DocxTemplate("template/template.docx")

# Data
context = {
    "employees": [
        { "name": "Pierre", "age": 22, "job": "IT guy", "gender": True },
        { "name": "Claire", "age": 31, "job": "HR ressources", "gender": False },
        { "name": "Gabriel", "age": 18, "job": "Dev intern", "gender": True },
        { "name": "Maxime", "age": 36, "job": "Secretary", "gender": True },
        { "name": "Zoé", "age": 27, "job": "Manager", "gender": False },
        { "name": "Alice", "age": 29, "job": "Designer", "gender": False },
        { "name": "Bob", "age": 24, "job": "Marketing", "gender": True },
        { "name": "Charlie", "age": 30, "job": "Sales", "gender": True },
        { "name": "Diana", "age": 28, "job": "Finance", "gender": False },
        { "name": "Ethan", "age": 35, "job": "Operations", "gender": True },
        { "name": "Fiona", "age": 26, "job": "Support", "gender": False },
        { "name": "George", "age": 33, "job": "Research", "gender": True },
        { "name": "Hannah", "age": 32, "job": "Development", "gender": False },
        { "name": "Ian", "age": 25, "job": "Quality Assurance", "gender": True },
        { "name": "Julia", "age": 34, "job": "Product Management", "gender": False },
        { "name": "Kevin", "age": 23, "job": "Data Analyst", "gender": True },
        { "name": "Laura", "age": 30, "job": "Customer Service", "gender": False },
        { "name": "Mike", "age": 29, "job": "Logistics", "gender": True },
        { "name": "Nina", "age": 27, "job": "Content Creation", "gender": False },
        { "name": "Oscar", "age": 31, "job": "Business Development", "gender": True },
        { "name": "Paula", "age": 28, "job": "Legal", "gender": False },
        { "name": "Quentin", "age": 24, "job": "IT Support", "gender": True },
        { "name": "Rachel", "age": 35, "job": "Public Relations", "gender": False },
        { "name": "Sam", "age": 26, "job": "Event Management", "gender": True },
        { "name": "Tina", "age": 32, "job": "Training and Development", "gender": False },
        { "name": "Ursula", "age": 30, "job": "Compliance", "gender": False },
        { "name": "Victor", "age": 29, "job": "Risk Management", "gender": True },
        { "name": "Wendy", "age": 34, "job": "Procurement", "gender": False },
        { "name": "Xander", "age": 23, "job": "IT Security", "gender": True },
        { "name": "Yara", "age": 31, "job": "Business Analysis", "gender": False },
        { "name": "Zane", "age": 28, "job": "Strategy", "gender": True }
    ],
    "version": "1.0",
    "introduction": "This is an introduction to python-docx-template.",
    "specificValue": "SPECIFIC_VALUE",
    "sections": [
        { "title": "First section", "subtitle": "First subtitle", "paragraph": "This is the first paragraph of my first section." },
        { "title": "Second section", "subtitle": "Second subtitle", "paragraph": "This is the first paragraph of my second section." },
        { "title": "Third section", "subtitle": "Third subtitle", "paragraph": "This is the first paragraph of my third section." },
    ],
    "dynamicImage": InlineImage(template, 'template/rte_logo.png', width=Mm(60), height=Mm(60))
}


# Replace img
template.replace_pic('template_img.png','./template/schema_unifilaire.png')

# testing that it works also when autoescape has been forced to True
jinja_env = jinja2.Environment(autoescape=True)

# Render the template with the context data
template.render(context)

# Save the generated document to a new file
template.save("report_generated.docx")