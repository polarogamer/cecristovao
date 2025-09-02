import pyautogui
import openpyxl
import pyperclip
from time import sleep

# Entrar na planilha
wordbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = wordbook['Produtos']
# - Copiar informaçao de um campo e colar no seu campo Correspodente
for linha in sheet_produtos.iter_rows(min_row=2):
        
        # Nome do Produto
        nome_produto = linha[0].value
        pyperclip.copy(nome_produto)
        pyautogui.click(771,223,duration=1)
        pyautogui.hotkey('ctrl','v')

        # Descrição
        descricao = linha[1].value
        pyperclip.copy(descricao)
        pyautogui.click(768,295,duration=1)
        pyautogui.hotkey('ctrl','v')
        
        # Categoria
        categoria = linha[2].value
        pyperclip.copy(categoria)
        pyautogui.click(767,409,duration=1)
        pyautogui.hotkey('ctrl','v')
    
        # Codigo Produto
        codigo_produto = linha[3].value
        pyperclip.copy(codigo_produto)
        pyautogui.click(769,481,duration=1)
        pyautogui.hotkey('ctrl', 'v')

        # Peso
        peso = linha[4].value
        pyperclip.copy(peso)
        pyautogui.click(770,562,duration=1)
        pyautogui.hotkey('ctrl','v')

        # Dimensões
        dimensoes = linha[5].value
        pyperclip.copy(dimensoes)
        pyautogui.click(768,638,duration=1)
        pyautogui.hotkey('ctrl','v')

        #pyautogui.moveTo(1271,249,duration=1)
        #pyautogui.scroll(-500)

        # Botão próximo, Pausa pra proxima Página
        pyautogui.click(969,681,duration=1)
        sleep(4)

        # Preço
        preco = linha[6].value
        pyperclip.copy(preco)
        pyautogui.click(778,241,duration=1)
        pyautogui.hotkey('ctrl','v')

        # Quantidade em estoque
        quantidade_em_estoque = linha[7].value
        pyperclip.copy(quantidade_em_estoque)
        pyautogui.click(776,314,duration=1)
        pyautogui.hotkey('ctrl','v')
        
        # Data de validade
        data_de_validade = linha[8].value
        pyperclip.copy(data_de_validade)
        pyautogui.click(775,391,duration=1)
        pyautogui.hotkey('ctrl','v')

        # Cor
        cor = linha[9].value
        pyperclip.copy(cor)
        pyautogui.click(776,468,duration=1)
        pyautogui.hotkey('ctrl','v')

        # Tamanho
        tamanho = linha[10].value
        pyautogui.click(778,540,duration=1)
        # Se for "Pequeno", clicar e uma pos
        if tamanho == 'Pequeno':
            sleep(0.1)
            pyautogui.click(797,597,duration=3)
        # Se for "Médio", clicar e uma posição
        elif tamanho == 'Médio':
            sleep(0.1)
            pyautogui.click(809,621,duration=3)
        # Se for "Grande", clicar e uma posição
        else:
            pyautogui.click(841,647,duration=3)


        # material 
        material = linha[11].value
        pyperclip.copy(material)
        pyautogui.click(776,616,duration=1)
        pyautogui.hotkey('ctrl','v')

        #pyautogui.moveTo(1342,401,duration=1)
        #pyautogui.scroll(-500)

        # Botão próximo, Pausa pra proxima Página
        pyautogui.click(950,663,duration=1)
        sleep(4)

        # Fabricante
        fabricante = linha[12].value
        pyperclip.copy(fabricante)
        pyautogui.click(776,258,duration=1)
        pyautogui.hotkey('ctrl','v')

        # País
        pais_origem = linha[13].value
        pyperclip.copy(pais_origem)
        pyautogui.click(776,336,duration=1)
        pyautogui.hotkey('ctrl','v')

        # Observacoes
        observacoes = linha[14].value
        pyperclip.copy(observacoes)
        pyautogui.click(776,406,duration=1)
        pyautogui.hotkey('ctrl','v')

        # codigo_de_barras
        codigo_de_barras = linha[15].value
        pyperclip.copy(codigo_de_barras)
        pyautogui.click(776,522,duration=1)
        pyautogui.hotkey('ctrl','v')

        # Localização_armazem
        localizacao_armazem = linha[16].value
        pyperclip.copy(localizacao_armazem)
        pyautogui.click(775,596,duration=1)
        pyautogui.hotkey('ctrl','v')

        # - Clicar em Proximo 
        pyautogui.click(954,645,duration=1)
        sleep(3)

        # - Clicar Botão Confirmar Salvo no Banco de Dados!
        pyautogui.click(1191,188,duration=1)
        sleep(3)

        # - Clicar no Botão Salvo no Banco de Dados!
        pyautogui.click(1199,189,duration=1)
        sleep(1)

        # Clicar no Botão em Adicionar mais um 
        pyautogui.click(1022,497,duration=1)
        sleep(2)


        

