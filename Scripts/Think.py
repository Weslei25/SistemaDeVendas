# Esses são os frameworks importados para criaço do sistema
# temos o QTdesigner e também o mysql.conector para conexo com o MySQL
from pycep_correios import get_address_from_cep, WebService, exceptions
from PyQt5 import QtWidgets, uic
import psycopg2
import xlwt
import datetime as dt
import pandas as pd
import win32com.client as win32
from pywintypes import com_error



banco = psycopg2.connect(
    host="fdefenderserver.ddns.net",
    user="postgres",
    password="(adm8081)",
    database="Weslei",
)


try:
    cur = banco.cursor()
    cur.execute("LOCK TABLE vendas IN ACCESS EXCLUSIVE MODE NOWAIT")
except psycopg2.OperationalError as e:
    if e.pgcode == psycopg2.errorcodes.LOCK_NOT_AVAILABLE:
        locked = True
    else:
        raise



# banco = mysql.connector.Connect(open('Config.txt', 'r').read().strip())



# Nome da função reponsavél por validar o login do usuario.
def chama_segunda_tela():
    # variavíes usadas na validação do usuario em questão.
    primeira_tela.label_4.setText("")
    nome_usuario = primeira_tela.lineEdit.text()
    senha = primeira_tela.lineEdit_2.text()
    # O curosr recebendo a conexão com o Banco de dados
    try:
        cursor = banco.cursor()
        # É passado as instruções sql que serão executadas nesse caso um select
        cursor.execute("""SELECT senha FROM
            usuarios WHERE nome ='{}'""".format(nome_usuario))
        # A variável senha_bd recebe tudo que o cursor pegou
        senha_bd = cursor.fetchall()
        # Temos a condição para validar o usuario
        if not nome_usuario:
            primeira_tela.label_4.setText("Preencha o campo LOGIN!")
            return
        elif not senha:
            primeira_tela.label_4.setText("Preencha o campo SENHA!")
            return
        if senha == senha_bd[0][0]:
            primeira_tela.close()
            formulario.show()
        cursor = banco.cursor()
        # É passado as instruções sql que serão executadas nesse caso um select
        # nas colunas especificas para melhorar os resultados, passando um order by
        cursor.execute("""SELECT idproduto, descricao, preco,
         observacao, marca,
         referencia from produtos order by idproduto DESC limit 1200""")

        dados_lidos1 = cursor.fetchall()
        # O objeto formulario recebe os objetos da interface
        # tableWidget.setRowCount, e a A len()função retorna
        # o número de itens em um objeto.
        formulario.tableWidget.setRowCount(len(dados_lidos1))
        formulario.tableWidget.setColumnCount(6)
        # Ele gera uma lista de números, que geralmente é usada para iterar com
        # forloops. Existem muitos casos de uso. Freqüentemente,
        # você desejará usar isso quando quiser executar uma ação X várias vezes
        for i in range(0, len(dados_lidos1)):
            for j in range(0, 6):
                formulario.tableWidget.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(dados_lidos1[i][j])))
        # Aqui temos um else para caso aldo de errado na
        # digitação dos dados de login


        cursor = banco.cursor()
        cursor.execute("""select idvenda, nomecliente, idtipo_negociacao,
            idvendedor, data_venda, dat_venv_fatuura, nomeproduto, 
            quantproduto, precoproduto, descproduto, vlr_total, natvenda,
            vezesdeparcelas, observacao FROM vendas order by idvenda DESC limit 1000000;""")

       
        sql_vendas1 = cursor.fetchall()
       
        formulario.tableWidget_4.setRowCount(len(sql_vendas1))
        formulario.tableWidget_4.setColumnCount(13)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 13):
                formulario.tableWidget_4.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

    except IndexError as indexx:
        aviso.show()
        aviso.textBrowser.setText("Login ou Senha incorretos\n\n{}".format(indexx))
        return
    else:
        primeira_tela.label_4.setText("Dados de login incorretos!")
        return
# Fechando o cursor com o metodo close()
    cursor.close()
# tela de sair clicando no botão sair essa função é executada

# Função responsavél por inserir os dados da tela de
# cadastro de produtos no banco
def cadastrar_produtos():
    # As variavés dos que recebém os inputs do usuario
    estoque = (formulario.lineEdit.text())
    descricao = str(formulario.lineEdit_2.text())
    preco = (formulario.lineEdit_3.text())
    # Recebe o valor da referencia do produto
    ref = (formulario.lineEdit_4.text())
    observacao = str(formulario.lineEdit_9.text())
    marca = str(formulario.lineEdit_5.text())
    categoria = str(formulario.comboBox_8.currentText())

    # Aqui temos algumas condições para verificar o valor
    # armazenado na variavél categoria e se esse valor estiver de acordo
    # com a condição ele recebe um outro valor inteiro para evitar
    # redundancias no banco de dados e conseqêntemente uso desnescessario
    # de espaço em disco
    if categoria == "Alimentos":
        categoria = "1"

    elif categoria == "Teste":
        categoria = "2"

    elif categoria == "Uso interno":
        categoria = "3"

    elif categoria == "Perfumaria":
        categoria = "4"

    elif categoria == "Roupa":
        categoria = "5"
    # se o valor não estiver de arcordo com nenhuma condção anterior
    #  ele recebe 5 por padão
    else:
        categoria = "5"

    try:
    
        cursor = banco.cursor()
        # Aqui uma variavél recebe os parametros sql, por que mais
        # adiante outra variavél é ultilizada para armazenar valores
        # que estão interligados
        SQL_produtos = """INSERT INTO produtos
         (idcategoria, descricao, preco, observacao,marca,referencia)
         VALUES  (%s,%s,%s,%s,%s,%s)"""
        # Esta variávell recebe por parametro as
        # outras variáveis que recebem os inputs do usuario
        dados1 = (categoria), (descricao), (preco), (observacao), (marca), (ref)
        # Esse cursor exeuta as duas variáveis em ordem como esta agora.
        if not preco:
            aviso.show()
            aviso.textBrowser.setText("  Preencha os campos vazios EX: Preço.")
            return
        elif not descricao:
            aviso.show()
            aviso.textBrowser.setText("  Preencha os campos vazios EX:Descrição.")
            return
        cursor.execute(SQL_produtos, dados1)
        # Como se trata de dados sendo insridos no BD temos
        # que usar o Metodo commit()
        banco.commit()
        
        cursor.close()
        # Após fazer todo o processo tenho que fechar com o metodo
        # close()
        cursor = banco.cursor()
        if not estoque:
            formulario.lineEdit_7.setText("Preencha o campo Estoque")
            return
        # O marcio usou procedures sql por isso esse metodo callproc()
        # foi usado baixo, e dentro dele é passad o nome da procedures sql
        # E a tabela que ela faz referencia
        cursor.callproc("PRC_EST", [estoque])
        # Como se trata de dados sendo insridos no BD temos
        # que usar o Metodo commit()
        banco.commit()
        # Novamente o uso do metodo cursor é nescessario
        # e o sql é passado dentro do metodo execute()
        cursor = banco.cursor()
        cursor.execute("""SELECT idproduto, descricao, preco, observacao, marca,
         referencia from produtos order by idproduto DESC limit 1200""")
        # temos aqui também o metodo fatchall() que pega todos os dados,
        # nesse caso eu armazeno esses dados em uma variavél
        sql_tprodu = cursor.fetchall()
        # O objeto formulario recebe os objetos da interface
        # tableWidget.setRowCount, e a A len()função retorna
        # o número de itens em um objeto.
        formulario.tableWidget.setRowCount(len(sql_tprodu))
        formulario.tableWidget.setColumnCount(6)

        for i in range(0, len(sql_tprodu)):
            for j in range(0, 6):
                formulario.tableWidget.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_tprodu[i][j])))
        # Acima é usado um for com a função range
        # e fechamos o cursor no fim
        cursor.close()
        # Após o usuario inserir os dados usamos o metodo setText("")
        # para limpar as linhas de input dando assim mais agilidade ao usuario
        formulario.lineEdit.setText("")
        formulario.lineEdit_2.setText("")
        formulario.lineEdit_3.setText("")
        formulario.lineEdit_4.setText("")
        formulario.lineEdit_5.setText("")
        formulario.lineEdit_9.setText("")

    except NameError as name:
        aviso.show()
        aviso.textBrowser.setText("Login ou Senha incorretos\n\n{}".format(name))
        return

# Função para  cadastrar novo usuario
def cadastrar_usuario():
    # As variavíes que estão iteradas com os lineedites
    # elas recebem os imputs com os dados que para facilitar
    # ja levam o nome do valor que o usuario precisa digitar
    nome = tela_cadastro.lineEdit.text()
    email = tela_cadastro.lineEdit_2.text()
    senha = tela_cadastro.lineEdit_3.text()
    c_senha = tela_cadastro.lineEdit_4.text()

    # Aqui temos uma condição e se essa condição for verdadeira
    # temos uma função que vai executar todo o tratamento da
    # criação do novo usuario
    if (senha == c_senha):
        try:

            cursor = banco.cursor()
            # O dados do usuario são armaznados em uma variável
            sql_user = """INSERT INTO usuarios (nome, email, senha)
            VALUES ('{}','{}','{}')""".format(nome, email, senha)
            # dados_user = str(nome), str(email), str(senha)
            # Os dados são executados
            cursor.execute(sql_user, dados_user)
            # Os dados são inseridos com o metodo commite()
            banco.commit()
            # As lineEdites são limpas com o metodo settext("")
            # Para criação de outro usuario se assim for nescessário
            nome = tela_cadastro.lineEdit.setText("")
            email = tela_cadastro.lineEdit_2.setText("")
            senha = tela_cadastro.lineEdit_3.setText("")
            c_senha = tela_cadastro.lineEdit_4.setText("")
            # O objeto tela cadastro recebe uma label que retornará a mensagem
            # abaixo caso o processo seja concluiido com sucesso
            tela_cadastro.label_2.setText("Usuario cadastrado com sucesso")
            cursor.close()
            # As except para tratamento do possivés erros
        except NameError as erro:
            tela_cadastro.label_2.setText('{}'.format(erro))

        except IndexError as erro2:
            tela_cadastro.label_2.setText('{}'.format(erro2))

        except ValueError as erro3:
            tela_cadastro.label_2.setText('{}'.format(erro3))

        except AttributeError as erro4:
           tela_cadastro.label_2.setText('{}'.format(erro4))

    # Um else caso somente a senha se estiver errada
    else:
        tela_cadastro.label_2.setText("As senhas digitadas estão diferentes")


def virificacep():

    try:
        cep = formulario.cepCliente.text()

        if not cep:
            aviso.show()
            aviso.textBrowser.setText("Preencha o campo CEP")
            return
        else:
            
            aviso.textBrowser.setText("Sucesso na consulta do CEP.")
            

        endereco = get_address_from_cep(
            cep, webservice=WebService.APICEP)

    except exceptions.CEPNotFound as testa:
        aviso.show()
        aviso.textBrowser.setText("Invalido,\n{}".format(testa))
        return

    except exceptions.ConnectionError as testaaa:
        aviso.show()
        aviso.textBrowser.setText("A conexão falhou,\n{}".format(testaaa))
        return

    except exceptions.Timeout as tempo:
        aviso.show()
        aviso.textBrowser.setText("Tempo excedido\n{}".format(tempo))
        return

    except exceptions.HTTPError as erro:
        aviso.show()
        aviso.textBrowser.setText("Ocorreu um erro\n{}".format(erro))
        return

    except exceptions.BaseException as base:
        aviso.show()
        aviso.textBrowser.setText(   "  CEP invalido. {}". format(base))
        
        return

    except ValueError as erru:
        aviso.show()
        aviso.textBrowser.setText('Valor não aceito\n\n{}'.format(erru))

        if not endereco:
            aviso.show()
            aviso.textBrowser.setText("atribuido!!!")
            return
        if not cep:
            aviso.show()
            aviso.textBrowser.setText("Cep não atribuido!!!")
        else:
            aviso.show()
            aviso.textBrowser.setText("Cep consultado com sucesso!!!")

    formulario.enderecoCliente.setText(endereco['logradouro'])
    formulario.bairroCliente.setText(endereco['bairro'])
    formulario.cidadeCliente.setText(endereco['cidade'])
    formulario.compleCliente.setText(endereco['logradouro'])
    formulario.lineEdit_8.setText(endereco['uf'])
    print(endereco)


def cadcliente():
    # Variaveis para cadastro de clientes
    nomedocliente = str(formulario.nomeCliente.text())
    cepdocliente = str(formulario.cepCliente.text())
    cidadedocliente = str(formulario.cidadeCliente.text())
    bairrodocliente = str(formulario.bairroCliente.text())
    ruadocliente = str(formulario.enderecoCliente.text())
    numerodocliente = str(formulario.numeroCliente.text())
    complemento = str(formulario.compleCliente.text())
    estadodocliente = str(formulario.lineEdit_8.text())
    celulardocliente = str(formulario.telCell.text())
    telefoneresidencial = str(formulario.telResid.text())
    categoriadocliente = str(formulario.catCliente.currentText())
    cpfdocliente = str(formulario.cpfCliente.text())
    rgdocliente = str(formulario.rgCliente.text())
    sitedocliente = str(formulario.siteCliente.text())
    emaildocliente = str(formulario.lineEdit_50.text())

    # Convertendo os valores da variavel "estadosdocliente" que recebe o uf dos estados
    try:
        if estadodocliente == "AC":
            estadodocliente = "1"
        if estadodocliente == "AL":
            estadodocliente = "2"
        if estadodocliente == "AM":
            estadodocliente = "3"
        if estadodocliente == "AP":
            estadodocliente = "4"
        if estadodocliente == "BA":
            estadodocliente = "5"
        if estadodocliente == "CE":
            estadodocliente = "6"
        if estadodocliente == "DF":
            estadodocliente = "7"
        if estadodocliente == "ES":
            estadodocliente = "8"
        if estadodocliente == "GO":
            estadodocliente = "9"
        if estadodocliente == "MA":
            estadodocliente = "10"
        if estadodocliente == "MG":
            estadodocliente = "11"
        if estadodocliente == "MS":
            estadodocliente = "12"
        if estadodocliente == "MT":
            estadodocliente = "13"
        if estadodocliente == "PA":
            estadodocliente = "14"
        if estadodocliente == "PB":
            estadodocliente = "15"
        if estadodocliente == "PE":
            estadodocliente = "16"
        if estadodocliente == "PI":
            estadodocliente = "17"
        if estadodocliente == "PR":
            estadodocliente = "18"
        if estadodocliente == "RJ":
            estadodocliente = "19"
        if estadodocliente == "RN":
            estadodocliente = "20"
        if estadodocliente == "RO":
            estadodocliente = "21"
        if estadodocliente == "RR":
            estadodocliente = "22"
        if estadodocliente == "RS":
            estadodocliente = "23"
        if estadodocliente == "SC":
            estadodocliente = "24"
        if estadodocliente == "SE":
            estadodocliente = "25"
        if estadodocliente == "SP":
            estadodocliente = "26"
        if estadodocliente == "TO":
            estadodocliente = "27"
        else:
            estadodocliente = "8"
        
    except ValueError as valorerrado:
        aviso.show()
        aviso.textBrowser.setText(" Algo deu errado.\n\n{}".format(valorerrado))
# As condições para inserir os dados do cliente no banco#
    try:

        if not cidadedocliente:
            aviso.show()
            aviso.textBrowser.setText("Preencha os campos com os dados solicitados!")
            return
        if cidadedocliente == 'Serra':
            cidadedocliente = int(70)

        elif cidadedocliente == 'Vila Velha':
            cidadedocliente = int(77)

        elif cidadedocliente == 'Vitória':
            cidadedocliente = int(78)

        elif cidadedocliente == 'Cariacica':
            cidadedocliente = int(17)

        if not bairrodocliente:
            aviso.show()
            aviso.textBrowser.setText("Informe o nome da bairro.")
            return

        elif  bairrodocliente == 'Feu Rosa':
            bairrodocliente = int(366)

        
        

        if not estadodocliente:
            aviso.show()
            aviso.textBrowser.setText("Informe o estado.")
            return

        elif not sitedocliente:
            sitedocliente = 'NULL'

        elif not emaildocliente:
            emaildocliente = 'NULL'

        if not nomedocliente:
            aviso.show()
            aviso.textBrowser.setText("Informe o nome do cliente.")
            return

        if not numerodocliente:
            aviso.show()
            aviso.textBrowser.setText("Informe o numero da residencia.")
            return
        
        if not celulardocliente:
            aviso.show()
            aviso.textBrowser.setText("Informe um numero de Celular.")
            return

        if not cpfdocliente:
            aviso.show()
            aviso.textBrowser.setText("Informe o CPF do cliente.")
            return


        if categoriadocliente == 'OCASIONAL':
            categoriadocliente = 'OCA'
        
        elif categoriadocliente == 'CLIENTE EXTRA':
            categoriadocliente = 'EXT'
        
        else:
            categoriadocliente = 'VIP'    
        
        
        bairrodocliente = int(366)
        cursor = banco.cursor()

        sql_cliente = """ INSERT INTO parceiros (idcidade, idestado,
        idbairro, nomeparc, cpf_cnpj, tipo_pessoa, cliente,
        cep, rua,numero, complemento, rg, tel_principal,
        tel_secund, email, site) VALUES(
            {}, {}, {}, '{}', '{}', 'F', '{}', '{}', '{}', '{}',
            '{}', '{}', '{}', '{}', '{}', '{}') """.format(
            cidadedocliente, estadodocliente,
            bairrodocliente,
            nomedocliente,
            cpfdocliente,
            categoriadocliente,
            cepdocliente,
            ruadocliente,
            numerodocliente,
            complemento,
            rgdocliente,
            celulardocliente,
            telefoneresidencial,
            emaildocliente,
            sitedocliente)

        cursor.execute(sql_cliente)
        
        banco.commit()
        cursor.close()
        aviso.show()
        aviso.textBrowser.setText("Cliente cadastrado com sucesso!!!")
    except IndexError as erro:
        aviso.show()
        aviso.textBrowser.setText("Alguns campos obrigratórios não foram preenchidos ou não é aceito o valor  inserido.{}".format(erro))
        return


def vender_produto():
    # Variaveis da tela de vendas
    datadavenda = dt.datetime.now()
    iddovendedor = formulario.comboBox_2.currentText()
    tipoNegociacao = formulario.tipoNegociacao.currentText()
    naturezavenda = formulario.naturezadavenda.currentText()
    vezesparcelas = formulario.vezesdeparcelas.currentText()
    datadevencimento = formulario.vencimentoparcelado.currentText()
    nomeclientevenda = formulario.lineEdit_14.text()
    nomeprodutovenda = formulario.lineEdit_15.text()
    quantidadedoprodutovenda = formulario.lineEdit_16.text()
    precodoproduto = formulario.lineEdit_17.text()
    descontodoproduto = formulario.lineEdit_18.text()
    valortotaldavenda = formulario.lineEdit_19.text()
    obcervacaovenda = formulario.lineEdit_6.text()
    idusuario = (3)


    if iddovendedor == "Alessandra":
        iddovendedor = '3'

    elif iddovendedor == "Eduarda":
        iddovendedor = '10'

    elif iddovendedor == "Jodeil":
        iddovendedor = '9'

    elif iddovendedor == "Luiz":
        iddovendedor = '7'
    else:
        iddovendedor = '3'
    # Aqui acontece a converssão de valores do tipo negociação
    if tipoNegociacao == "Distributiva":
        tipoNegociacao = "15"

    elif tipoNegociacao == "Integrativa":
        tipoNegociacao = "16"

    elif tipoNegociacao == "Adversarial":
        tipoNegociacao = "17"

    elif tipoNegociacao == "Cooperativa ou colaborativa":
        tipoNegociacao = "18"

    elif tipoNegociacao == "Direta":
        tipoNegociacao = "19"

    elif tipoNegociacao == "Indireta":
        tipoNegociacao = "20"

    elif tipoNegociacao == "Ganha-Ganha":
        tipoNegociacao = "21"

    elif tipoNegociacao == "Perde-Perde":
        tipoNegociacao = "22" 

    elif tipoNegociacao == "Autonegociação":
        tipoNegociacao = "23"

    # Vezes parceladas
    if vezesparcelas == '1X':
        vezesparcelas = '1'

    elif vezesparcelas == '2X':
        vezesparcelas = '2'

    elif vezesparcelas == '3X':
        vezesparcelas = '3'

    elif vezesparcelas == '4X':
        vezesparcelas = '4'

    elif vezesparcelas == '5X':
        vezesparcelas = '5'

    elif vezesparcelas == '6X':
        vezesparcelas = '6'

    elif vezesparcelas == '7X':
        vezesparcelas = '7'

    elif vezesparcelas == '8X':
        vezesparcelas = '8'

    elif vezesparcelas == '9X':
        vezesparcelas =  '9'

    elif vezesparcelas == '10X':
        vezesparcelas = '10'

    if not nomeclientevenda:
        aviso.show()
        aviso.textBrowser.setText("  Informe o NOME ou CPF do cliente.")
        return
    elif not nomeprodutovenda:
        aviso.show()
        aviso.textBrowser.setText("  Declare o nome do produto.")
        return


    try:
        
        cursor = banco.cursor()

        sql_vendas = """ INSERT INTO vendas (idtipo_negociacao, idusuarios,
        idvendedor, data_venda, vlr_total, nomecliente, nomeproduto,quantproduto, precoproduto,descproduto, natvenda, vezesdeparcelas, observacao, dat_venv_fatuura)
        VALUES({}, {}, {}, '{}',{}, '{}', '{}', {}, {}, {}, '{}', {}, '{}', {});""".format(
            tipoNegociacao, idusuario,
            iddovendedor, datadavenda,valortotaldavenda, nomeclientevenda, nomeprodutovenda, 
            quantidadedoprodutovenda, precodoproduto, descontodoproduto, naturezavenda,
             vezesparcelas, obcervacaovenda, datadevencimento)



        
        cursor.execute(sql_vendas)
        banco.commit()
        
        
        cursor = banco.cursor()
        cursor.execute("""select idvenda, nomecliente, 
            idtipo_negociacao, idvendedor, data_venda, dat_venv_fatuura, nomeproduto, 
            quantproduto, precoproduto, descproduto, vlr_total, natvenda,
            vezesdeparcelas, observacao FROM vendas order by idvenda DESC limit 1000000;""")

       
        sql_vendas1 = cursor.fetchall()
       
        formulario.tableWidget_4.setRowCount(len(sql_vendas1))
        formulario.tableWidget_4.setColumnCount(13)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 13):
                formulario.tableWidget_4.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

        formulario.lineEdit_19.setText("")
        formulario.lineEdit_10.setText("")
        formulario.lineEdit_17.setText("")
        formulario.lineEdit_14.setText("")
        formulario.lineEdit_15.setText("")


    except IndexError as erros:
        aviso.show()
        aviso.textBrowser.setText("Preencha os campos obrigatorios.\n\n{}".format(erros))
        return


    except ValueError as erro:
        aviso.show()
        aviso.textBrowser.setText("Formato de dados invalido.\n\n{}".format(erro))
        return
    except psycopg2.errors.InFailedSqlTransaction as er:
        aviso.show()
        aviso.textBrowser.setText("Algo deu errado.\n\n{}".format(er))
        return
    except psycopg2.errors.SyntaxError as erros:
        aviso.show()
        aviso.textBrowser.setText("Algo deu errado.\n\n{}".format(erros))


# função para deletar um produto


def deletar_produto():
    
    try:
        # A variavél cursor recebe a conexão com o BD e recebe o metodo cursor()
        # para ser executada mais afrente com o cursor.execute
        cursor = banco.cursor()
        deletarvenda = formulario.lineEdit_11.text()
        # Variavél que recebe o sql com metodo .format e dentro do format recebe
        # o valor da variavl que recebe o valor do lineEdite, assim dando pra
        # usar o valor que o usuario digitar
        # sql_del = """DELETE FROM estoque WHERE idproduto = {}""".format(cod_del)
        # Tratando-se de uma tabela que tem uma chave estrangeira atrelada
        # a ela precisa de um tratamento para excluir os dados
        deletarvendasql = """DELETE FROM vendas WHERE idvenda = {}""".format(deletarvenda)
        # Dentro dos parenteses passamos as variáveis que sql_del, sql_del2
        # para executar junto ao banco de dados o sql desejado
        if not deletarvenda:
            aviso.show()
        aviso.textBrowser.setText('É nescessario declarar o codigo da venda a dar baixa.')
        return 

        cursor.execute(deletarvendasql)
        
        aviso.show()
        aviso.textBrowser.setText('Produto deletado com exito.')
    except NameError as erro:
        aviso.show()
        aviso.textBrowser.setText('Preencha os campos corretamente.\n\n{}'.format(erro))
    except psycopg2.errors.SyntaxError as er:
        aviso.show()
        aviso.textBrowser.setText('Preencha os campos corretamente.\n\n{}'.format(er))
    except psycopg2.errors.InFailedSqlTransaction as erros:
        aviso.show()
        aviso.textBrowser.setText('Algo deu errodo.\n\n{}'.format(erros))


# função que abre a tela de cadastro
def abre_tela_cadastro():
    # o metodo show() é usado para chamar a tela
    tela_cadastro.show()

def gerarelatorio():
    df = pd.read_sql('select * from Vendas', banco)
    df.to_excel("C:\\Users\\Public\\Desktop\\Relatorio_Vendas.xlsx", index=False)
    tela_progresso.show()
    tela_progresso.progressBar.setValue(100)



def gerarelatorio_produtos():
    df = pd.read_sql('select * from produtos', banco)

    df.to_excel("C:\\Users\\Public\\Desktop\\Relatorio_Produtos.xlsx", index=False)
    tela_progresso.show()
    tela_progresso.progressBar.setValue(100)

def enviaremail():

    try:
        # criar integração com outlook

        outlook = win32.Dispatch('outlook.application')

        # criar um email de envio 
        email = outlook.CreateItem(0)

        # configurar as informações do email
        faturamento = str(formulario.lineEdit_25.text())
        vendas = str(formulario.lineEdit_26.text())
        ticket = str(formulario.lineEdit_27.text())

        email.To = str(formulario.lineEdit_13.text())
        email2 = email.to
        email.Subject = str(formulario.lineEdit_22.text())
        email.HTMLBody = """

        <p>Olá bom dia!</p>

        <p>Aqui é da loja 1 o faturamento da loja foi de {}.</p>

        <p>Vendemos {} produtos.</p>

        <p>O ticket medio foi de {}.</p>

        <p>Abraçoes loja 1.</p>

        """.format(faturamento, vendas, ticket)
        if not email.to:
            aviso.show()
            aviso.textBrowser.setText('Microsoft Outlook Precisamos saber para quem enviar isto. Verifique se voc� inseriu pelo menos um nome.')
            return
        email.Send()
        aviso.show()
        aviso.textBrowser.setText('Email enviado para\n{}'.format(email2))

        formulario.lineEdit_25.setText('')
        formulario.lineEdit_26.setText('')
        formulario.lineEdit_27.setText('')
        formulario.lineEdit_13.setText('')
        formulario.lineEdit_22.setText('')
    except NameError as erro:
        aviso.show()
        aviso.textBrowser.setText('Algo deu errado \n\n{}'.format(erro))
        return
    except ConnectionError as erro:
        aviso.show()
        aviso.textBrowser.setText('Algo deu errado \n\n{}'.format(erro))
        return
    except com_error as erros:
        aviso.show()
        aviso.textBrowser.setText('Algo deu errado \n\n{}'.format(erros))
        return


def enviaremailprodutos():

    try:
        # criar integração com outlook
        # email =  think_V1@outlook.com
        # senha = weslei080319
        outlook = win32.Dispatch('outlook.application')

        # criar um email de envio 
        email = outlook.CreateItem(0)

        # configurar as informações do email
        faturamento = str(formulario.lineEdit_32.text())
        vendas = str(formulario.lineEdit_33.text())
        ticket = str(formulario.lineEdit_34.text())
        cdialogo = str(formulario.lineEdit_7.text())

        email.To = str(formulario.lineEdit_29.text())
        email2 = email.to
        email.Subject = str(formulario.lineEdit_31.text())
        email.HTMLBody = """

        <p>Olá bom dia!</p>

        <p>Aqui é da loja 1 o faturamento da loja foi de {}.</p>

        <p>Vendemos {} produtos.</p>

        <p>O ticket medio foi de {}.</p>

        <p>{}.</p>

        <p>Abraçoes loja 1.</p>

        """.format(faturamento, vendas, ticket, cdialogo)
        if not email.to:
            aviso.show()
            aviso.textBrowser.setText('Microsoft Outlook Precisamos saber para quem enviar isto. Verifique se voc� inseriu pelo menos um nome.')
            return
        email.Send()
        aviso.show()
        aviso.textBrowser.setText('Email enviado para\n{}'.format(email2))

        formulario.lineEdit_25.setText('')
        formulario.lineEdit_26.setText('')
        formulario.lineEdit_27.setText('')
        formulario.lineEdit_13.setText('')
        formulario.lineEdit_22.setText('')
    except NameError as erro:
        aviso.show()
        aviso.textBrowser.setText('Algo deu errado \n\n{}'.format(erro))
        return
    except ConnectionError as erro:
        aviso.show()
        aviso.textBrowser.setText('Algo deu errado \n\n{}'.format(erro))
        return
    except com_error as erros:
        aviso.show()
        aviso.textBrowser.setText('Algo deu errado \n\n{}'.format(erros))
        return

def vendas_realizadas():
    try:
        cursor = banco.cursor()
        cursor.execute("""select idvenda, nomecliente, idtipo_negociacao,
                idvendedor, data_venda, dat_venv_fatuura, nomeproduto, 
                quantproduto, precoproduto, descproduto, vlr_total, natvenda,
                vezesdeparcelas, observacao FROM vendas order by idvenda DESC limit 1000000;""")


        sql_vendas1 = cursor.fetchall()
           
        formulario.tableWidget_5.setRowCount(len(sql_vendas1))
        formulario.tableWidget_5.setColumnCount(14)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 14):
                formulario.tableWidget_5.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

        formulario.tableWidget_6.setRowCount(len(sql_vendas1))
        formulario.tableWidget_6.setColumnCount(14)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 14):
                formulario.tableWidget_6.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

        cursor.close()
    except psycopg2.errors.InFailedSqlTransaction as erro:
        aviso.show()
        aviso.textBrowser.setText('Algo deu errado \n\n{}'.format(erro))


# Conectores e operadores
# App recebe os objetos QT
app = QtWidgets.QApplication([])
app.setStyle ( 'Windows' )
# As variavés abaixo recebem e carregam as telas com o metodo loadUi
primeira_tela = uic.loadUi("teladelogin.ui")
tela_cadastro = uic.loadUi("tela_cadastro.ui")
formulario = uic.loadUi("segunda_tela.ui")
aviso = uic.loadUi("avisosnovos.ui")
tela_progresso = uic.loadUi("barradeprogreço.ui")
# Abaixo conectamos um objeto ao outro para que possão interagir conforme
# nosso projeto, ex: A variavél primeira_tela recebe o botão e o botão
# recebe o metodo clicKed com o metodo connect para conectar a função
# que é passada dentro como parametro
formulario.salvarCliente.clicked.connect(cadcliente)
primeira_tela.pushButton.clicked.connect(chama_segunda_tela)
formulario.pushButton_9.clicked.connect(deletar_produto)
formulario.enviaemail.clicked.connect(enviaremail)
formulario.actionRELATORIO_DE_VENDAS.triggered.connect(gerarelatorio)

# Aqui usamos o metodo Password para indicar que esse campo vai receber
# um valor do tipo senha e recebera o '*' como padão de segurança.
primeira_tela.lineEdit_2.setEchoMode(QtWidgets.QLineEdit.Password)
# Os botões que Abrem as telas
formulario.pushButton_8.clicked.connect(virificacep)
primeira_tela.pushButton_2.clicked.connect(abre_tela_cadastro)
tela_cadastro.pushButton.clicked.connect(cadastrar_usuario)
formulario.pushButton.clicked.connect(cadastrar_produtos)
formulario.pushButton_4.clicked.connect(vender_produto)
formulario.pushButton_10.clicked.connect(vendas_realizadas)
formulario.geraexcel.clicked.connect(gerarelatorio)
formulario.geraexcel_2.clicked.connect(gerarelatorio_produtos)
formulario.enviaemail_2.clicked.connect(enviaremailprodutos)
formulario.pushButton_12.clicked.connect(vendas_realizadas)
# Essa é uma lista que o combobox recebe pra puxar os dados
# que estão no codigo, passos esses parametros adicionando os itens
# que quero que apareça no combobox


formulario.comboBox_2.addItems(
    ["Alessandra", "Eduarda", "Jodeil", "Luiz"])
formulario.tipoNegociacao.addItems(
    ['Distributiva',
'Integrativa',
'Adversarial',
'Cooperativa ou colaborativa',
'Direta',
'Indireta',
'Ganha-Ganha',
'Perde-Perde',
'Autonegociação'])

formulario.comboBox_8.addItems(
    ["Roupa", "Teste", "Uso interno",
     "Perfumaria", "Alimentos"])

formulario.naturezadavenda.addItems(
    ["Dinheiro", "Credito", "Debito", "Crediario", "Cheque"])

formulario.catCliente.addItems(
    ["VIP", "OCASIONAL", "CLIENTE EXTRA"])

formulario.vencimentoparcelado.addItems([ '1', '2', '3', '4', '5',
'6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18',
'19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30'  ])
formulario.vezesdeparcelas.addItems(['1X',
    '2X', '3X','4X', '5X', '6X', '7X', '8X', '9X', '10X'
    ])
# Esse premeiro show() é aonde o sistema começa.
primeira_tela.show()
# formulario.show()

app.exec()
