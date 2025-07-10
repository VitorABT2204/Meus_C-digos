import pandas as pd
import numpy as np
import openpyxl as op
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

arq=pd.read_excel('TabelaExemplo.xlsx') # abrir arquivo excel
Tab=arq #variavel contendo só as áreas principais da tabela

while(True):
 
    def grafico(): #função que da um gráfico dos itens mais vendidos
        cel = 0
        lap =0
        per =0
        mon = 0
        for i  in range(len(Tab)): 
        #loop que le na tabela se a linha designada na coluna Categoria é de uma categoria expecifica e contabiliza a quantidade de itens das categorias
            if Tab.loc[i,"Categoria"]=='Celulares':#o Tab.loc[i,"Categoria"]=='Celulares' verifica se na linha i da tabela a coluna Categoria é igua a Celulares
                cel += 1
            elif  Tab.loc[i,"Categoria"]=='Laptop':#o Tab.loc[i,"Categoria"]=='Celulares' verifica se na linha i da tabela a coluna Categoria é igua a Laptop
                lap += 1
            elif  Tab.loc[i,"Categoria"]=='Periféricos':#o Tab.loc[i,"Categoria"]=='Celulares' verifica se na linha i da tabela a coluna Categoria é igua a Periféricos
               per +=1
            elif  Tab.loc[i,"Categoria"]=='Monitor':#o Tab.loc[i,"Categoria"]=='Celulares' verifica se na linha i da tabela a coluna Categoria é igua a Monitor
                mon +=1
        
        plt.ylabel('Unidades')
        y = np.array([cel, lap, per, mon]) #cria um array com as quantidades de itens das categorias
        x = np.array(["Celulares", "Laptop", "Periféricos", "Monitor"]) #cria um array com os nomes das categorias
        
        plt.bar(x,y, color=['b','r','yellow','g']) #cria o gráfico com o eixo X e Y e da às suas 4 barras cores distintas
        
        plt.title('Itens mais vendidos') # da um título pro gráfico
        plt.xlabel('Itens') #escreve um label em baixo do eixo X
        plt.xticks(rotation=45) #rotaciona os textos do eixo X em 45 graus 
        plt.ylabel('Quantidades') #escreve um label no lado do eixo Y
        #criação do gráfico de barras
        
        return plt.show()
    
    def lucro(cat, ano):
        val_por_ano = {} #cria um dicionário vazio que terá como chave o ano e seu valor será uma lista de valores
    
        # Agrupar os valores por ano
        for i in range(len(Tab)):
            if Tab.loc[i, "Categoria"] == cat and Tab.loc[i, "Ano"] <= ano: 
                #filtra os dados da tabela mantendo apenas aqueles que possuem a mesma categoria contida em cat e que o ano seja menor ou igual a variável ano
                
                ano_atual = int(Tab.loc[i, "Ano"]) #pega o ano da linha atual e o transforma em inteiro (caso o ano seja uma string)
                valor = Tab.loc[i, "Valor"] #pega o valor da linha atual
                
                if ano_atual not in val_por_ano:
                    val_por_ano[ano_atual] = [valor]
                    #esse if verifica se esse ano ainda não existe no dicionário e cria uma nova lista com esse valor.
                else:
                    val_por_ano[ano_atual].append(valor)
                    #Mas caso o ano já exista no dicionário, adiciona o valor à lista existente.
    
        # Calcular a média dos valores por ano
        anos = sorted(val_por_ano.keys()) #Cria uma lista com os anos do dicionário e ordena cronologicamente.
        medias = [np.mean(val_por_ano[ano]) for ano in anos] #É um list comprehension que calcula a média dos valores de cada ano
    
        # Criar o gráfico
        plt.figure(figsize=(8,5)) #Cria uma nova janela do matplotlib com tamanho desejado (8x5 no caso)
        plt.bar(anos, medias, color='b') #Cria um gráfico onde o eixo X é a lista ano e o eixo Y é as médias e as barras tem a cor azul
        plt.gca().xaxis.set_major_locator(MaxNLocator(integer=True))  # força o eixo x a ter só números inteiros
        plt.xlabel("Ano") #Cria um texto em baixo do eixo X
        plt.ylabel("Média de Preço") #Cria um texto no lado do eixo Y
        plt.title(f"Média de Lucro para '{cat}' até {ano}") #Dá um título para o gráfico
        
        return plt.show()
    
    def grafico_reg():
       
        agrupado = Tab.groupby(['Região', 'Categoria'])['Valor'].sum().reset_index()
        #O 'Tab.groupby(['Região', 'Categoria']):' agrupa os dados por região e categoria.
        # Já '['Valor'].sum():' soma os valoresdentro de cada grupo.
        # E o '.reset_index():' transforma o resultado em um DataFrame normal, trazendo Região e Categoria de volta como colunas (e não como índice).

        
        mais_vendidas = agrupado.loc[agrupado.groupby('Região')['Valor'].idxmax()]
        #Nesse o 'groupby('Região')['Valor']:' agrupa o agrupado por região, olhando só para os valores.
        #O 'idxmax():' retorna o índice da linha com o maior valor em cada região.

        mais_vendidas['Categoria_Regiao'] = mais_vendidas['Categoria'] + ' - ' + mais_vendidas['Região']
        #Cria uma nova coluna Categoria_Regiao juntando os nomes da categoria e da região, separado-os por " - ".

        # "Print"
        plt.figure(figsize=(10,6)) #Cria uma nova janela do matplotlib com tamanho desejado (10x6 no caso)
        plt.bar(mais_vendidas['Categoria_Regiao'], mais_vendidas['Valor'], color='green') 
        #cria o gráfico onde o eixo X mostra os nomes de Categoria_Regiao, o eixo Y é o valor total vendido e a cor das barras é verde
        
        plt.xlabel('Categoria - Região') #Cria um texto em baixo do eixo X
        plt.ylabel('Total Vendido') #Cria um texto no lado do eixo Y
        plt.title('Categoria mais vendida por Região') #Dá um título para o gráfico
        
        plt.xticks(rotation=45) #rotaciona os textos do eixo x em 45 graus
        plt.tight_layout() #Ajusta o layout automaticamente para evitar que textos fiquem cortados nas bordas.
        
        return plt.show()


    escolha = input('\nVocê deseja ver a tabela toda ou só uma de suas categorias: 1(Tabela toda), 2(Categoria): ') 
    #parte onde o usário escolhe se quer a tabela toda ou apena uma de suas categorias

    if(escolha == '1'):
        print(arq) #print da tabela toda

    elif(escolha=='2'):
        categoria = input(" Qual Categoria de produtos você deseja acessar: 1(Celulares), 2(Laptop), 3(Periféricos), 4(Monitor): ") 
        #escolha de qual categoria o usuário deseja ver

        match categoria:
            case '1':
                print(Tab[Tab['Categoria']=='Celulares']) #print da categoria Celulares
                categoria = 'Celulares'

            case '2':
                print(Tab[Tab['Categoria']=='Laptop']) #print da categoria Laptops
                categoria = 'Laptop'
        
            case '3':
                print(Tab[Tab['Categoria']=='Periféricos']) #print da categoria 'Periféricos'
                categoria = 'Periféricos'
                
            case '4':
                print(Tab[Tab['Categoria']=='Monitor']) #print da categoria 'Monitor'
                categoria = 'Monitor'
                
            case '_':
                print('ERROR: escolha inválida')#print de Erro caso o usuario escolha uma opção diferente de 1(Celulares), 2(Laptop), 3(Periféricos), 4(Monitor)
                break   #sai do loop e encerra o programa
            
        ver_gra = input('\nDeseja ver o grafico das categorias para ver qual foi a mais vendida S ou N: ') 
        #pergunta se o usuário deseja ver o gráfico que mostra a categoria com mais vendas
        
        if ver_gra.upper() == 'S':
             grafico()
        elif ver_gra.upper() == 'N':
            print('Próximo')
        else:
            print('Error opção invalida')
            break #sai do loop e encerra o programa
            
        ver_lucro = input(f'\nDeseja ver o grafico do lucro da categoria {categoria} ao longo dos anos: S ou N: ') 
        #pergunta se o usuário deseja ver o gráfico que o lucro ao logo dos anos
        if ver_lucro.upper() == 'S':
            anos = int(input('\nAté qual ano você deseja ver (2019 - 2021): '))
            #pergunta qual o intervalo de tempo desejado
            if anos <= 2021 and anos >=2019:
                lucro(categoria,anos)
            else:
                print('Fora do intevalo')

        ver_cat_reg = input('\nDeseja ver um gráfico que mostre a categoria mais vendida por região S ou N: ')
        if ver_cat_reg.upper() == 'S':
             grafico_reg()
        elif ver_cat_reg.upper() == 'N':
            print('Próximo')
        else:
            print('Error opção invalida')
            break #sai do loop e encerra o programa
            
    else:
        print('ERROR: escolha inválida')#print de Erro caso o usuario escolha uma opção diferente de 1(Tabela toda), 2(Categoria)
        break #sai do loop e encerra o programa
    
    continua = input('\nContinuar para outra verificação: S ou N: ')

    if continua.upper() == 'S':
        continue #continua no loop pro usuário escolher outra opção
    elif continua.upper() == 'N':
        break #sai do loop e encerra o programa
    else:
        print('Error opção de continuação invalida')
