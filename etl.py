import pandas as pd
import openpyxl
import openai


#    extraçao

repositorio_df = pd.read_excel("Pasta.xlsx")

Voluntario = repositorio_df["nome"].tolist()
hobbie = repositorio_df["hobbie"] 
materia_fav = repositorio_df["materia preferida"] is ["name"]

print(Voluntario)


#   transformação
openai_api_key = "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"
openai.api_key = openai_api_key

def gerador_mensagens(user):
    
    completion = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
      {
          "role": "system",
          "content": "Você é um especialista em exames vocacionais."
      },
      {
          "role": "user",
          "content": f"Crie uma mensagem para {Voluntario} dando dicas de qual profissão seguir de acordo com sua materia escolar preferida:{materia_fav} e seu hobbie preferido: {hobbie}com um máximo de 1 palavra"
      }
    ]
  )
    return completion.choices[0].message.content.strip('\"')



msg = gerador_mensagens(repositorio_df)
print(msg)

#  carregamento

novo_repositorio = openpyxl.load_workbook("Pasta.xlsx")
page = novo_repositorio["Sheet1"]


page.append(["mensagem",str(msg)])
        

novo_repositorio.save("Pasta.xlsx")
print(repositorio_df)

print(repositorio_df)



