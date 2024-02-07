from docxtpl import DocxTemplate

doc = DocxTemplate("invoice_template.docx")

invoice_list = [[2, 'pen', 0.5, 1],
                [1, 'paper', 5, 5],
                [2, 'notebook', 2, 4]]


doc.render({"name":"john",
            "phone": 55,
            "invoice_list": invoice_list,
            "Sub-total": 10,
            "salestax" : "5%",
            "total": 9
            })
doc.save("new_invoice.docx")