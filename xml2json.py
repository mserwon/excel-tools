import fitz
import json
import re

jtext = ''
type = 'Script'
odoc = fitz.open("OperatorV1-out.oxps")
for page in odoc:
   match type:
       case 'Script':
          jt = page.get_text("text")
          if re.search("^Name", jt):
              type = 'Variable'
              dvar = page.get_text("dict")
          else:
              jtext += jt
       case 'Variable':
           dv = page.get_text("dict")
           if dv['blocks'][0]['lines'][0]['spans'][0]['text'] == 'Resources':
               type = 'Resources'
               dres = page.get_text("dict")
           else:
               #dvar2 = page.get_text("dict")
               for block in dv['blocks']:
                   dvar['blocks'].append(block)
       case 'Resources':
           pass

json_data = json.dumps(dvar)
with open("uwm.json", "w") as json_file:
        json_file.write(json_data)