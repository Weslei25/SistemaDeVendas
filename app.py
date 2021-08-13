from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import * 
import mysql.connector
import win32com.client as win32
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PyQt5.QtGui import * 
from PyQt5.QtCore import *
from pycep_correios import get_address_from_cep, WebService, exceptions
import requests
import datetime as dt
import pandas as pd
import json

"""
vendedor
entradaOuSaida
clienteVenda
cnpjCpf
prosutoVenda
quantidadeVenda
cdDeBarras
descontoVEnda
PorcentoVenda
categoriaVenda
quantidadeitens
quantidadeProdutos
troco
saldoDevedor
totalDeDesconto
TotalVendas
"""
with open('Config\\config.json') as f:# Abrindo o arquivo que contem a string de conexao com o DB
    entrada = json.load(f)# lendo o json e armazenando em uma variavel

conexao = mysql.connector.Connect(# string de conexao
    host=entrada["host"],
    user=entrada["user"],
    password=entrada["password"],
    database=entrada["database"],
    auth_plugin=entrada["auth_plugin"]
)


def chama_segunda_tela():# função responsavel por chamar a tela principal
    
    telaDeLogin.label_4.setText("")# Sempre limpo o campo de avisos para que quando o usuario corrigir o erro o campo esteja limpo
    nome_usuario = telaDeLogin.lineEdit.text()# Aqui pego o nome de usuario
    senha = telaDeLogin.lineEdit_2.text()# Aqui pego a senha
    

    try:
        verificaUsuario = pd.read_sql(f"select nome from usuarios where nome='{nome_usuario}'", conexao)

        if verificaUsuario.empty == True:
            telaDeLogin.label_4.setText("Usuario não encontrado!")
            return
        else:
            verificaUsuario = (verificaUsuario['nome'][0])

        verificaSenha = pd.read_sql(f"select senha FROM usuarios WHERE nome ='{nome_usuario}'", conexao)
        if verificaSenha.empty == True:
            telaDeLogin.label_4.setText("Usuario não encontrado!")
            return
        else:
            verificaSenha = (verificaSenha['senha'][0])
        
        
        if not nome_usuario:
            telaDeLogin.label_4.setText("Preencha o campo LOGIN!")
            return
        elif not senha:
            telaDeLogin.label_4.setText("Preencha o campo SENHA!")
            return

        if senha == verificaSenha:
            telaDeLogin.close()
            TelaPrincipal.show()
        else:
            telaDeLogin.label_4.setText("Usuario ou Senha incorretos!")
        
        TelaPrincipal.usuarios.setText('Usuario Logado: {}'.format(nome_usuario))
        datahoje = dt.datetime.now()

        nomedousuario = pd.read_sql(f"select idusuarios from usuarios where nome='{nome_usuario}'", conexao)
        nomedousuario = (nomedousuario['idusuarios'][0])
        cursor = conexao.cursor()
        cursor.execute(f"INSERT INTO log_usuario(descricao, idusuarios, dt_logusuario) VALUES ('NULL', {nomedousuario}, '{datahoje}');")
        conexao.commit()

        
        def catalogarProdutos():
            
            cursor = conexao.cursor()
            cursor.execute("""SELECT idproduto, descricao, (DATE_FORMAT(dt_entrada , '%d/%m/%Y')), preco, observacao, marca,
            referencia from produtos order by idproduto DESC limit 1200""")

            dados_lidos1 = cursor.fetchall()
            TelaPrincipal.tableWidget.setRowCount(len(dados_lidos1))
            TelaPrincipal.tableWidget.setColumnCount(7)
            for i in range(0, len(dados_lidos1)):
                for j in range(0, 7):
                    TelaPrincipal.tableWidget.setItem(
                        i, j, QtWidgets.QTableWidgetItem(str(dados_lidos1[i][j])))

        catalogarProdutos()

        cursor = conexao.cursor()
        cursor.execute("""select idvenda, nomecliente,
            tipo_negociacao_idtipo_negociacao,
            vendedores_idvendedor, (DATE_FORMAT(data_venda , '%d/%m/%Y')), dat_venv_fatuura, nomeproduto,
            quantproduto, precoproduto, descproduto, vlr_total, id_tipopagamento,
            vezesdeparcelas, observacao, entrada_saida FROM vendas order by idvenda DESC limit 1000000""")


        sql_vendas1 = cursor.fetchall()

        TelaPrincipal.tableWidget_5.setRowCount(len(sql_vendas1))
        TelaPrincipal.tableWidget_5.setColumnCount(15)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 15):
                TelaPrincipal.tableWidget_5.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

        # Estou pegando a data do dia nessa variavél datavenda
        datadehoje = dt.datetime.now()
        # Estou setando o valor da variavél no objeto da data
        #TelaPrincipal.dateEdit_4.setDate(datadehoje)
        TelaPrincipal.dateEdit_5.setDate(datadehoje)
        # TelaPrincipal.dateEdit_2.setDate(datadehoje)
        TelaPrincipal.dateEdit_4.setDate(datadehoje)
        TelaPrincipal.dateEdit.setDate(datadehoje)
        TelaPrincipal.dateEdit.setDate(datadehoje)
        # TelaPrincipal.dateEdit_8.setDate(datadavenda)
        # TelaPrincipal.dateEdit_9.setDate(datadavenda)
        # TelaPrincipal.dateEdit_10.setDate(datadavenda)
        # TelaPrincipal.dateEdit_14.setDate(datadavenda)
        # TelaPrincipal.dateEdit_15.setDate(datadavenda)
        # TelaPrincipal.dateEdit_16.setDate(datadavenda)

        cursor.close()
        
    except Exception as indexx:
        telaDeLogin.label_4.setText("{}".format(indexx))
        return
    else:
        telaDeLogin.label_4.setText("Dados de login incorretos!")
        return
    






def virificacep():

    try:
        cep = TelaPrincipal.cepCliente.text()

        if not cep:
            return False
        else:
            aviso.textBrowser.setText("Sucesso na consulta do CEP.")


        endereco = get_address_from_cep(cep, webservice=WebService.APICEP)

    except exceptions.CEPNotFound as notfound:
        aviso.show()
        aviso.textBrowser.setText("{}".format(notfound))
        return

    except exceptions.ConnectionError as testaaa:
        aviso.show()
        aviso.textBrowser.setText("{}".format(testaaa))
        return

    except exceptions.Timeout as tempo:
        aviso.show()
        aviso.textBrowser.setText("{}".format(tempo))
        return

    except exceptions.HTTPError as erro:
        aviso.show()
        aviso.textBrowser.setText("{}".format(erro))
        return

    except exceptions.BaseException as base:
        aviso.show()
        aviso.textBrowser.setText("{}". format(base))

        return

    except ValueError as erru:
        aviso.show()
        aviso.textBrowser.setText('Valor não aceito\n\n{}'.format(erru))

    TelaPrincipal.enderecoCliente.setText(endereco['logradouro'])
    TelaPrincipal.bairroCliente.setText(endereco['bairro'])
    TelaPrincipal.cidadeCliente.setText(endereco['cidade'])
    TelaPrincipal.compleCliente.setText(endereco['logradouro'])
    TelaPrincipal.lineEdit_8.setText(endereco['uf'])


def cadcliente():
    try:
        # Variaveis para cadastro de clientes
        nomedocliente = str(TelaPrincipal.nomeCliente.text())
        cepdocliente = str(TelaPrincipal.cepCliente.text())
        cidadedocliente = str(TelaPrincipal.cidadeCliente.text())
        bairrodocliente = str(TelaPrincipal.bairroCliente.text())
        ruadocliente = str(TelaPrincipal.enderecoCliente.text())
        numerodocliente = str(TelaPrincipal.numeroCliente.text())
        complemento = str(TelaPrincipal.compleCliente.text())
        estadodocliente = str(TelaPrincipal.lineEdit_8.text())
        celulardocliente = str(TelaPrincipal.telCell.text())
        telefoneresidencial = str(TelaPrincipal.telResid.text())
        categoriadocliente = str(TelaPrincipal.catCliente.currentText())
        cpfdocliente = str(TelaPrincipal.cpfCliente.text())
        rgdocliente = str(TelaPrincipal.rgCliente.text())
        sitedocliente = str(TelaPrincipal.siteCliente.text())
        emaildocliente = str(TelaPrincipal.lineEdit_50.text())

        verificaCpfExiste = pd.read_sql(f"""select cpf_cnpj from parceiros p where cpf_cnpj ='{cpfdocliente}';""", conexao)
        verificaCpfExiste = (verificaCpfExiste['cpf_cnpj'][0])
        if verificaCpfExiste == cpfdocliente:
            aviso.show()
            aviso.textBrowser.setText("Cliente ja cadastrado!")
            return
        else:
            aviso.show()
            aviso.textBrowser.setText("Entrei aqui")
        

        cursor = conexao.cursor()

        df = pd.read_sql(f"SELECT idestado from estados where sigla = '{estadodocliente}'", conexao)
        estadodocliente = (df['idestado'][0])
        

        df = pd.read_sql(f"SELECT idcidade FROM cidades WHERE nome = '{cidadedocliente}'", conexao)
        cidadedocliente = (df['idcidade'][0])
        

        if not bairrodocliente:
            aviso.show()
            aviso.textBrowser.setText("Informe o nome da bairro.")
            return

        
        df = pd.read_sql(f"SELECT idbairro FROM bairros WHERE nome = '{bairrodocliente}'", conexao)
        bairrodocliente = (df['idbairro'][0])
        print(bairrodocliente)
        
        
        if not estadodocliente:
            aviso.show()
            aviso.textBrowser.setText("Informe o estado.")
            return
        if not cidadedocliente:
            aviso.show()
            aviso.textBrowser.setText("Preencha os campos com os dados solicitados!")
            return
        if not sitedocliente:
            sitedocliente = 'NULL'
            
        if not emaildocliente:
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


        sql_cliente = """ insert into parceiros (bairros_cidades_idcidade, bairros_cidades_estados_idestado,
        bairros_idbairro, nomeparc, cpf_cnpj, tipo_pessoa, cliente,
        cep, rua,numero, complemento, rg, tel_principal,
        tel_secund, email, site) values(
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
        conexao.commit()
        cursor.close()

        aviso.show()
        aviso.textBrowser.setText("Cliente cadastrado com sucesso!!!")
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return




def geraRelatorioVendasEntSaida():
    try:

        datavenda1 = TelaPrincipal.dateEdit_4.text()
        datavenda1 = dt.datetime.strptime(datavenda1, '%d/%m/%Y')
        datavenda1 = datavenda1.strftime('%Y-%m-%d')

        datavenda2 = TelaPrincipal.dateEdit_5.text()
        datavenda2 = dt.datetime.strptime(datavenda2, '%d/%m/%Y')
        datavenda2 = datavenda2.strftime('%Y-%m-%d')
        ent_sai = TelaPrincipal.comboBox_6.currentText()

        if ent_sai == '1 - Entrada':
            ent_sai = int(1)
        elif ent_sai == '2 - Saida':
            ent_sai =  int(2)
        


        df = pd.read_sql("""select idvenda, tipo_negociacao_idtipo_negociacao, vendedores_usuarios_idusuarios, vendedores_idvendedor, (DATE_FORMAT(data_venda , '%d/%m/%Y')), vlr_total, nomecliente, nomeproduto,
            quantproduto, precoproduto, descproduto, id_tipopagamento, vezesdeparcelas,
            observacao, dat_venv_fatuura, entrada_saida
            FROM vendas where entrada_saida = {} and data_venda >= '{}' and data_venda <= '{}'order
            by idvenda DESC limit 1000000;""".format(ent_sai,datavenda1, datavenda2), conexao)
        
        # Mudando o nome das colunas que seram geradas no excel
        df = df.rename(columns={'idvenda': 'Codigo Da Venda', 'tipo_negociacao_idtipo_negociacao':'Tipo De Negociação', 'vendedores_usuarios_idusuarios':'Usuario','vendedores_idvendedor':'Vendedor', "(DATE_FORMAT(data_venda , '%d/%m/%Y'))":'Data Da Venda','vlr_total':'Valor Total', 'nomecliente':'Nome Do Cliente', 'nomeproduto':'Nome Do Produto',
        'quantproduto':'Quantidade Vendida', 'precoproduto':'Preço Do Produto', 'descproduto':'Desconto','id_tipopagamento':'Tipo De Pagamento', 'vezesdeparcelas':'Vezes De Parcelas', 'observacao':'Observação', 'dat_venv_fatuura':'Data De Cobrança', 'entrada_saida':'Entrda ou Saida'})

        salvar = QtWidgets.QFileDialog.getSaveFileName()[0]
        df.to_excel(salvar + '.xlsx', index=False)

        tela_progresso.show()
        tela_progresso.progressBar.setValue(100)
    except ValueError as er:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(er))
        
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))

def gerarrelatorioprodutos():
    
    try:

        datavenda1 = TelaPrincipal.dateEdit_8.text()
        datavenda1 = dt.datetime.strptime(datavenda1, '%d/%m/%Y')
        datavenda1 = datavenda1.strftime('%Y-%m-%d')

        datavenda2 = TelaPrincipal.dateEdit_9.text()
        datavenda2 = dt.datetime.strptime(datavenda2, '%d/%m/%Y')
        datavenda2 = datavenda2.strftime('%Y-%m-%d')
        


        df = pd.read_sql("""SELECT idproduto, descricao, (DATE_FORMAT(dt_entrada , '%d/%m/%Y')), preco, observacao, marca,
            referencia from produtos where dt_entrada >='{}' and dt_entrada <='{}' order by idproduto DESC limit 1200""".format(datavenda1, datavenda2), conexao)
        
        # Renomeando as colunas 
        # df.rename(columns={'$a':'a', '$b':'b', '$c':'c', '$d':'d', '$e':'e'})
        df = df.rename(columns={'idproduto':'Codigo Do Produto', 'descricao':'Descrição Do Produto', "(DATE_FORMAT(dt_entrada , '%d/%m/%Y'))":'Data Do Cadastro', 'preco':'Preço', 'observacao':'Obsercação Do Produto', 'marca':'Fabricante'})
        

        salvar = QtWidgets.QFileDialog.getSaveFileName()[0]
        df.to_excel(salvar + '.xlsx', index=False, )

        tela_progresso.show()
        tela_progresso.progressBar.setValue(100)
    except ValueError as er:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(er))
        
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))

def fecharbarradeprogreco():
    
    tela_progresso.close()
    

def vendasAvista():
    # Aqui é feito a converssão das datas
    categoriaDavenda = TelaPrincipal.comboBox_6.currentText()

    datavenda1 = TelaPrincipal.dateEdit_4.text()
    datavenda1 = dt.datetime.strptime(datavenda1, '%d/%m/%Y')
    datavenda1 = datavenda1.strftime('%Y-%m-%d')

    datavenda2 = TelaPrincipal.dateEdit_5.text()
    datavenda2 = dt.datetime.strptime(datavenda2, '%d/%m/%Y')
    datavenda2 = datavenda2.strftime('%Y-%m-%d')

    if categoriaDavenda == ('Dinheiro'):
        categoriaDavenda = int(6)
    elif categoriaDavenda == ('Credito'):
        categoriaDavenda = int(7)
    elif categoriaDavenda ==  ('Debito'):
        categoriaDavenda = int(8)
    elif categoriaDavenda == ('Crediario'):
        categoriaDavenda = int(9)
    elif  categoriaDavenda == ('Cheque'):
        categoriaDavenda = int(10)
    else:
        categoriaDavenda = int(1)


    try:
        cursor =  conexao.cursor()


        cursor.execute(""" select idvenda, nomecliente,
                tipo_negociacao_idtipo_negociacao,
                vendedores_idvendedor, (DATE_FORMAT(data_venda , '%d/%m/%Y')), dat_venv_fatuura, nomeproduto,
                quantproduto, precoproduto, descproduto, vlr_total, id_tipopagamento,
                vezesdeparcelas, observacao, entrada_saida FROM vendas where vezesdeparcelas <= 1
                and data_venda >= '{}' and data_venda <= '{}' order by idvenda DESC limit 1000000;""".format(datavenda1, datavenda2))
        sqlVendasAvista = cursor.fetchall()

        TelaPrincipal.tableWidget_5.setRowCount(len(sqlVendasAvista))

        TelaPrincipal.tableWidget_5.setColumnCount(15)

        for i in range(0, len(sqlVendasAvista)):
            for j in range(0, 15):
                TelaPrincipal.tableWidget_5.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVendasAvista[i][j])))
        cursor.close()
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return

def enviaemail():

    telaDeEmail.show()
    pass



def realizarvendas():
    
    telaDeVendas.show()

    pegaridusuario = pd.read_sql("select idusuarios from log_usuario where idusuarios is not null order by id_logusuario DESC limit 1", conexao)
    pegaridusuario = (pegaridusuario['idusuarios'][0])

    df = pd.read_sql(f"select nome from usuarios where idusuarios='{pegaridusuario}'", conexao)
    pegaridusuario = str(df['nome'][0])

    telaDeVendas.nomevendedor.setText(pegaridusuario)

def fecharteladeaviso():
    aviso.close()

def vender_produto():

    try:
    	# dthoje = dt.datetime.now()
        # Variaveis da tela de vendas
        datadavenda = dt.datetime.now()
        iddovendedor = TelaPrincipal.vendedor_2.currentText()
        tipoNegociacao = TelaPrincipal.tipoNegociacao.currentText()
        naturezavenda = TelaPrincipal.naturezadavenda.currentText()
        vezesparcelas = TelaPrincipal.vezesdeparcelas.currentText()
        datadevencimento = TelaPrincipal.vencimentoparcelado.currentText()
        nomeclientevenda = TelaPrincipal.lineEdit_14.text()
        nomeprodutovenda = TelaPrincipal.lineEdit_15.text()
        quantidadedoprodutovenda = TelaPrincipal.lineEdit_16.text()
        precodoproduto = TelaPrincipal.lineEdit_17.text()
        descontodoproduto = TelaPrincipal.lineEdit_18.text()
        valortotaldavenda = TelaPrincipal.lineEdit_19.text()
        obcervacaovenda = TelaPrincipal.lineEdit_6.text()
        entrada_saida = TelaPrincipal.comboBox_3.currentText()
        idusuario = (3)


        df = pd.read_sql(f"select idvendedor from vendedores where nome ='{iddovendedor}'", conexao)
        iddovendedor = (df['idvendedor'][0])
        # Aqui acontece a converssão de valores do tipo negociação

        df = pd.read_sql(f"select idtipo_negociacao from tipo_negociacao where descricao ='{tipoNegociacao}'", conexao)
        tipoNegociacao = (df['idtipo_negociacao'][0])
        
        
        df = pd.read_sql(f"select id_entrada_saida from etrada_saida where descricao ='{entrada_saida}'", conexao)
        entrada_saida = (df['idtipo_negociacao'][0])
        

        if naturezavenda == ('Dinheiro'):
            naturezavenda = int(6)
        elif naturezavenda == ('Credito'):
            naturezavenda = int(7)
        elif naturezavenda ==  ('Debito'):
            naturezavenda = int(8)
        elif naturezavenda == ('Crediario'):
            naturezavenda = int(9)
        elif  naturezavenda == ('Cheque'):
            naturezavenda = int(10)
        else:
            return False

        if not nomeclientevenda:
            aviso.show()
            aviso.textBrowser.setText("  Informe o NOME ou CPF do cliente.")
            return
        elif not nomeprodutovenda:
            aviso.show()
            aviso.textBrowser.setText("  Declare o nome do produto.")
            return
        else:

            cursor = conexao.cursor()

            sql_vendas = """ INSERT INTO vendas (tipo_negociacao_idtipo_negociacao, vendedores_usuarios_idusuarios,
            vendedores_idvendedor, data_venda, vlr_total, nomecliente, nomeproduto,quantproduto, precoproduto,descproduto, id_tipopagamento, vezesdeparcelas, observacao, dat_venv_fatuura, entrada_saida)
            VALUES({}, {}, {}, '{}',{}, '{}', '{}', {}, {}, {}, '{}', {}, '{}', {}, {});""".format(
                tipoNegociacao, idusuario,
                iddovendedor, datadavenda,valortotaldavenda, nomeclientevenda, nomeprodutovenda,
                quantidadedoprodutovenda, precodoproduto, descontodoproduto, naturezavenda,
                vezesparcelas, obcervacaovenda, datadevencimento, entrada_saida)

            cursor.execute(sql_vendas)
            conexao.commit()


            cursor = conexao.cursor()
            cursor.execute("""select idvenda, nomecliente,
                tipo_negociacao_idtipo_negociacao,
                vendedores_idvendedor, data_venda, dat_venv_fatuura, nomeproduto,
                quantproduto, precoproduto, descproduto, vlr_total, id_tipopagamento,
                vezesdeparcelas, observacao, entrada_saida FROM vendas order by idvenda DESC limit 1000000;""")


            sql_vendas1 = cursor.fetchall()

            TelaPrincipal.tableWidget_4.setRowCount(len(sql_vendas1))
            TelaPrincipal.tableWidget_4.setColumnCount(15)

            for i in range(0, len(sql_vendas1)):
                for j in range(0, 15):
                    TelaPrincipal.tableWidget_4.setItem(
                        i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

            TelaPrincipal.lineEdit_19.setText("")
            # TelaPrincipal.lineEdit_10.setText("")
            TelaPrincipal.lineEdit_17.setText("")
            TelaPrincipal.lineEdit_14.setText("")
            TelaPrincipal.lineEdit_15.setText("")
            TelaPrincipal.lineEdit_16.setText("")

    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return
        

def cadastrar_produtos():

    try:

        datadaentrada = dt.datetime.now()
        
        estoque = (TelaPrincipal.estoque.text())
        descricao = str(TelaPrincipal.descricao.text())
        preco = (TelaPrincipal.preco.text())
        ref = (TelaPrincipal.referencia.text())
        observacao = str(TelaPrincipal.observacao.text())
        marca = str(TelaPrincipal.marca.text())
        categoria = str(TelaPrincipal.categotiaproduto.currentText())

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

        else:
            categoria = "5"



        cursor = conexao.cursor()
        SQL_produtos = """INSERT INTO produtos (categorias_idcategoria, descricao, preco, observacao,marca,referencia, dt_entrada)
         VALUES  ('{}', '{}', {}, '{}','{}', '{}', '{}')""".format(categoria, descricao, preco, observacao, marca, ref, datadaentrada)
        # dados1 = (categoria), (descricao), (preco), (observacao), (marca), (ref), (datadaentrada)
        if not preco:
            aviso.show()
            aviso.textBrowser.setText("  Preencha os campos vazios EX: Preço.")
            return
        elif not descricao:
            aviso.show()
            aviso.textBrowser.setText("  Preencha os campos vazios EX:Descrição.")
            return
        cursor.execute(SQL_produtos)
        conexao.commit()
    
        if not estoque:
            aviso.show()
            aviso.textBrowser.setText("Preencha o campo estoque")
            return

        cursor.execute("SELECT MAX(idproduto) FROM produtos")
        produtoult = cursor.fetchall()
        tratadoproduto = produtoult[0][0]
        cursor.execute("INSERT INTO estoque (estoque, produtos_idproduto) values ({},{})".format(estoque,tratadoproduto))
        conexao.commit()
        
        cursor.execute("""SELECT idproduto, descricao, (DATE_FORMAT(dt_entrada , '%d/%m/%Y')), preco, observacao, marca,
         referencia from produtos order by idproduto DESC limit 1200""")
        sql_tprodu = cursor.fetchall()

        TelaPrincipal.tableWidget.setRowCount(len(sql_tprodu))
        TelaPrincipal.tableWidget.setColumnCount(7)

        for i in range(0, len(sql_tprodu)):
            for j in range(0, 7):
                TelaPrincipal.tableWidget.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_tprodu[i][j])))

        cursor.close()

        """
        TelaPrincipal.lineEdit_2.setText("")
        TelaPrincipal.lineEdit_3.setText("")
        TelaPrincipal.lineEdit_4.setText("")
        TelaPrincipal.lineEdit_5.setText("")
        TelaPrincipal.lineEdit_9.setText("")
        """
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return

def vendas_parceladas():
    # Aqui é feito a converssão das datas
    datavenda1 = TelaPrincipal.dateEdit_4.text()
    datavenda1 = dt.datetime.strptime(datavenda1, '%d/%m/%Y')
    datavenda1 = datavenda1.strftime('%Y-%m-%d')

    datavenda2 = TelaPrincipal.dateEdit_5.text()
    datavenda2 = dt.datetime.strptime(datavenda2, '%d/%m/%Y')
    datavenda2 = datavenda2.strftime('%Y-%m-%d')

    try:

        cursor = conexao.cursor()
        cursor.execute("""select idvenda, nomecliente,
                tipo_negociacao_idtipo_negociacao,
                vendedores_idvendedor, (DATE_FORMAT(data_venda , '%d/%m/%Y')), dat_venv_fatuura, nomeproduto,
                quantproduto, precoproduto, descproduto, vlr_total, id_tipopagamento,
                vezesdeparcelas, observacao, entrada_saida FROM vendas
                where vezesdeparcelas > 1 and data_venda >= '{}' and data_venda <= '{}' order by idvenda DESC limit 1000000;""".format(datavenda1, datavenda2))


        sql_vendas1 = cursor.fetchall()

        TelaPrincipal.tableWidget_5.setRowCount(len(sql_vendas1))
        TelaPrincipal.tableWidget_5.setColumnCount(15)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 15):
                TelaPrincipal.tableWidget_5.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

        cursor.close()

    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return


def deletarProduto():
    try:
        codigoestoque = TelaPrincipal.codEstoque.text()
        codigoDoproduto = TelaPrincipal.codProduto_2.text()
        if not codigoestoque:
            aviso.show()
            aviso.textBrowser.setText('Preencha o campo Codigo/Estoque')
        elif not codigoDoproduto:
            aviso.show()
            aviso.textBrowser.setText('Preencha o campo Codigo/Produto')

        cursor = conexao.cursor()
        cursor.execute("""delete from estoque where produtos_idproduto ={}""".format(codigoestoque))
        conexao.commit()
        cursor.execute("""delete from produtos where idproduto={}""".format(codigoDoproduto))
        conexao.commit()
    
        
        cursor = conexao.cursor()
        cursor.execute("""SELECT idproduto, descricao, (DATE_FORMAT(dt_entrada , '%d/%m/%Y')), preco, observacao, marca,
        referencia from produtos order by idproduto DESC limit 1200""")

        dados_lidos1 = cursor.fetchall()
        TelaPrincipal.tableWidget.setRowCount(len(dados_lidos1))
        TelaPrincipal.tableWidget.setColumnCount(7)
        for i in range(0, len(dados_lidos1)):
            for j in range(0, 7):
                TelaPrincipal.tableWidget.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(dados_lidos1[i][j])))
    
        TelaPrincipal.codEstoque.setText("")
        TelaPrincipal.codProduto_2.setText("")

        aviso.show()
        aviso.textBrowser.setText('Produto codigo {} foi deletado com exito'.format(codigoDoproduto))

    except Exception as e:
        if not codigoestoque:
            aviso.show()
            aviso.textBrowser.setText('Preencha o campo Codigo/Estoque')
        elif not codigoDoproduto:
            aviso.show()
            aviso.textBrowser.setText('Preencha o campo Codigo/Produto')
        else:
            aviso.show()
            aviso.textBrowser.setText('{}'.format(e))
        return

       
def consultarcnpj():


    # https://receitaws.com.br/api

    
    try:
        cnpj = TelaPrincipal.cnpj_consulta.text()

        cnpj = ''.join(c for c in cnpj if c.isdigit())# retirar pontos e virgulas


        url = 'https://www.receitaws.com.br/v1/cnpj/{}'.format(cnpj)

        
        pegar_dados = requests.get(url)# Pegando os dados
        dados_obtidos = pegar_dados.json()# trasformando os dados em um json

        
        TelaPrincipal.nomeEmpresa.setText(dados_obtidos['nome'].upper())
        TelaPrincipal.tipo_cnpj.setText(dados_obtidos['tipo'].upper())
        TelaPrincipal.nomeFantazia.setText(dados_obtidos['fantasia'].upper())
       #  TelaPrincipal.dateEdit.setDate(dados_obtidos['abertura'].upper())
        TelaPrincipal.situacaoEmpresa.setText(dados_obtidos['situacao'].upper())
        TelaPrincipal.capitalSocialEmpresa.setText(dados_obtidos['capital_social'].upper())
        TelaPrincipal.naturezaJuridica.setText(dados_obtidos['natureza_juridica'].upper())
        TelaPrincipal.cepEmpresa.setText(dados_obtidos['cep'].upper())
        TelaPrincipal.municipioEmpresa.setText(dados_obtidos['municipio'].upper())
        TelaPrincipal.bairroEmpresa.setText(dados_obtidos['bairro'])
        TelaPrincipal.complementoEmpresa.setText(dados_obtidos['complemento'].upper())
        TelaPrincipal.numeroEnderecoEmpresa.setText(dados_obtidos['numero'].upper())
        TelaPrincipal.emailEmpresa.setText(dados_obtidos['email'].upper())
        TelaPrincipal.telefoneEmpresa.setText(dados_obtidos['telefone'].upper())
        TelaPrincipal.porteEmpresa.setText(dados_obtidos['porte'].upper())
        TelaPrincipal.ufEmpresa.setText(dados_obtidos['uf'].upper())
        TelaPrincipal.logradouroEmpresa.setText(dados_obtidos['logradouro'].upper())
        atividadePrincipal = dados_obtidos['atividade_principal']
        
        for atividade in atividadePrincipal:
            atividade = atividade['text']
        TelaPrincipal.atividadePrincipal.setText(atividade.upper())

        atividadeSecundaria = dados_obtidos['atividades_secundarias']
        for dado in atividadeSecundaria:
            dado = dado['text']
        TelaPrincipal.atividadeSecundarias.setText(dado.upper())
        # atividades_secundarias

    except Exception as erro:
        aviso.show()
        aviso.textBrowser.setText("{}". format(erro))
    finally:
        pass

def recuperasenhalogin():

    nomeusuario = telaDeLogin.lineEdit.text()
    
    retorno = pd.read_sql(f"select email, senha from usuarios where nome ='{nomeusuario}'", conexao)

    emaildb = (retorno['email'][0])
    
    senha = (retorno['senha'][0])
    
    try:
        outlook = win32.Dispatch('outlook.application')

        email = outlook.CreateItem(0)

        email.To = str(emaildb)
        email2 = email.to
        email.Subject = ('Recuperação De Senha Do Sistema De VEndas;')
        email.HTMLBody = """

        <p>Olá Tudo bem?</p>

        <p>A sua senha do sistema de vendas é {}.</p>

        <p>Att, .</p>

        <p>Sistema De Vendas em microempresas.</p>

        """.format(senha)
        if not email.to:
            aviso.show()
            aviso.textBrowser.setText('Email não localizado')
            return
        elif email == 'None':
            aviso.show()
            aviso.textBrowser.setText('Email não localizado')
        else:
            email.Send()
            aviso.show()
            aviso.textBrowser.setText('Sua senha foi enviada para o email {}'.format(email2))
                
    except Exception as e:
        if e == "(-2147352567, 'Exceção.', (4096, 'Microsoft Outlook', 'O Outlook não reconhece um ou mais nomes. ', None, 0, -2147467259), None)":
            e == 'Email não encontrado na nossa base de dados'
            
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return

def tela_cadastrousuario():

    tela_cadastro.show()

def cadastrar_usuario():

    nome = tela_cadastro.lineEdit.text()
    email = tela_cadastro.lineEdit_2.text()
    senha = tela_cadastro.lineEdit_3.text()
    c_senha = tela_cadastro.lineEdit_4.text()
    data_criacao = dt.datetime.now()

    if not nome:
        tela_cadastro.label_2.setText('Preencha os campos')
        return
    if (senha == c_senha):
        try:

            cursor = conexao.cursor()
            sql_user = """INSERT INTO usuarios (nome, email, senha, created_at)
            VALUES ('{}','{}','{}', '{}')""".format(nome, email, senha, data_criacao)
            cursor.execute(sql_user)
            conexao.commit()



            nome = tela_cadastro.lineEdit.setText("")
            email = tela_cadastro.lineEdit_2.setText("")
            senha = tela_cadastro.lineEdit_3.setText("")
            c_senha = tela_cadastro.lineEdit_4.setText("")

            tela_cadastro.label_2.setText("Usuario cadastrado com sucesso")
            cursor.close()

        except NameError as erro:
            tela_cadastro.label_2.setText('{}'.format(erro))
            return
        except IndexError as erro2:
            tela_cadastro.label_2.setText('{}'.format(erro2))
            return
        except ValueError as erro3:
            tela_cadastro.label_2.setText('{}'.format(erro3))
            return
        except AttributeError as erro4:
           tela_cadastro.label_2.setText('{}'.format(erro4))
           return
    # Um else caso somente a senha se estiver errada
    else:
        tela_cadastro.label_2.setText("As senhas digitadas estão diferentes")

    
def pesquisarProduto():
    try:
        pesquisar = TelaPrincipal.lineEdit_12.text()
        cursor = conexao.cursor()
        cursor.execute("""SELECT idproduto, descricao, (DATE_FORMAT(dt_entrada , '%d/%m/%Y')), preco, observacao, marca,
         referencia from produtos where descricao like '{}'""".format(pesquisar))

        sqlVerificacaoProduto = cursor.fetchall()
        TelaPrincipal.tableWidget.setRowCount(len(sqlVerificacaoProduto))
        TelaPrincipal.tableWidget.setColumnCount(7)

        for i in range(0, len(sqlVerificacaoProduto)):
            for j in range(0, 6):
                TelaPrincipal.tableWidget.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVerificacaoProduto[i][j])))

        cursor.close()
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return

def cadastrar_empresa():
    try:
        cnpj = TelaPrincipal.cnpj_consulta.text()
        razao_social = TelaPrincipal.nomeEmpresa.text()
        nome_fantazia = TelaPrincipal.nomeFantazia.text()
        tipo_empresa = TelaPrincipal.tipo_cnpj.text()
        atividade_principal = TelaPrincipal.atividadePrincipal.text()
        natureza_juridica = TelaPrincipal.naturezaJuridica.text()
        atividade_secundaria = TelaPrincipal.atividadeSecundarias.text()
        situacao = TelaPrincipal.situacaoEmpresa.text()
        capital_socia = TelaPrincipal.capitalSocialEmpresa.text()
        cep_empresa = TelaPrincipal.cepEmpresa.text()
        complemento = TelaPrincipal.complementoEmpresa.text()
        email_empresa = TelaPrincipal.emailEmpresa.text()
        telefone_empresa = TelaPrincipal.telefoneEmpresa.text()
        abertura_empresa = TelaPrincipal.dateEdit.text()
        porte_empresa = TelaPrincipal.porteEmpresa.text()
        idbairros = TelaPrincipal.bairroEmpresa.text()
        idcidade = TelaPrincipal.municipioEmpresa.text()
        idestado = TelaPrincipal.ufEmpresa.text()
        
        abertura_empresa = dt.datetime.strptime(abertura_empresa, '%d/%m/%Y')
        abertura_empresa = abertura_empresa.strftime('%Y-%m-%d')

        idcidade = pd.read_sql(f"select idcidade from cidades where nome ='{idcidade}'", conexao)
        if idcidade.empty == True:
            aviso.show()
            aviso.textBrowser.setText('Cidade não informada!')
            return
        else:
            idcidade = (idcidade['idcidade'][0])

        idestado = pd.read_sql(f"select idestado from estados where sigla ='{idestado}'", conexao)
        if idestado.empty == True:
            aviso.show()
            aviso.textBrowser.setText('Estado não informado!')
            return
        else:
            idestado = (idestado['idestado'][0])

        idbairro = pd.read_sql(f"select idbairro from bairros where nome ='{idbairros}'", conexao)
        if idbairro.empty == True:
            cursor =  conexao.cursor()
            cursor.execute("""INSERT INTO think.bairros (cidades_estados_idestado, cidades_idcidade, nome) VALUES ({}, {}, '{}');""".format(idestado, idcidade, idbairros))
            conexao.commit()
            cursor.close()
        else:
            idbairro = (idbairro['idbairro'][0])

        VerificaEmpresa = pd.read_sql(f"select cnpj from empresa where razao_social ='{razao_social}' and cnpj = '{cnpj}'", conexao)
        if VerificaEmpresa.empty == True:
            cursor = conexao.cursor()
            cursor.execute("""INSERT INTO think.empresa (cnpj, razao_social, nome_fantazia, tipo_empresa, atividade_principal, natureza_juridica, atividade_secundaria, situacao, capital_social, cep, complemento, email_empresa, telefone, abertura_empresa, porte_empresa, idbairro, idcidade, idestado)
                            VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', {}, '{}', '{}', '{}', '{}', '{}', '{}', {}, {}, {})""".format(cnpj, razao_social, nome_fantazia, tipo_empresa, atividade_principal, natureza_juridica, atividade_secundaria, situacao, capital_socia, cep_empresa, complemento, email_empresa, telefone_empresa, abertura_empresa, porte_empresa, idbairro, idcidade, idestado))
            conexao.commit()
            cursor.close()
            aviso.show()
            aviso.textBrowser.setText('Empresa cadastrada com sucesso!')
        
        else:
            VerificaEmpresa = (VerificaEmpresa['cnpj'][0])
            if VerificaEmpresa == cnpj:
                aviso.show()
                aviso.textBrowser.setText('Empresa ja cadastrada!')
            return
    except Exception as erro:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(erro))
   

def arquivoaserenviado():
   


    try:
        salvar = QtWidgets.QFileDialog.getOpenFileName()[0]# Pegando path co arquivo a ser enviado
        telaDeEmail.lineEdit.setText(salvar)# Setando o caminho em um lineEdit
    except Exception as erro:
        print("\n{}\n".format(erro))

def enviaremailcomarquivo():
    
    try:
        # 
        # email =  think_V1@outlook.com
        # senha = weslei080319
        # Email do sistema sistemadevendasecadastro2522@gmail.com
        
        emailDestinatario = telaDeEmail.lineEdit_2.text() # pego o destinatario 
        anexodoemail = telaDeEmail.lineEdit.text()# Recebo o anexo do email se tiver.
        corpoDoEmail = telaDeEmail.textEdit.toPlainText()# QTextEdit.toPlainText é a propriedade que aceita a quebra de linha no qtextEdit
        
        fromaddr = "sistemadevendasecadastro2522@gmail.com"# Email remetente
        toaddr = emailDestinatario # Email destinatario
        msg = MIMEMultipart()

        msg['From'] = fromaddr 
        msg['To'] = toaddr
        msg['Subject'] = telaDeEmail.lineEdit_3.text()

        body = corpoDoEmail + "\n\n\nAtt\n\nSistema de VEndas e Cadastro" # Corpo do email

        msg.attach(MIMEText(body, 'plain'))

        filename = anexodoemail # arquivo

        attachment = open(f'{anexodoemail}','rb') # abrir e verificar arquivo

        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())# Carregar arquivo
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)

        attachment.close()
        senha = entrada['senhaemail']# Pegando a senha do email

        server = smtplib.SMTP('smtp.gmail.com', 587)# Servidor de email
        server.starttls()# Criptografia de arquivo
        server.login(fromaddr, senha)# login no email, passando a senha
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)# Finalizando envio email
        server.quit()
        
        aviso.show()
        aviso.textBrowser.setText('Email enviado com sucesso para {}'.format(emailDestinatario))
        
    except Exception as erro:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(erro))
        return


app = QtWidgets.QApplication([]) # Criando um QApplication que faz a construcao da minha aplicacao.
app.setStyle ( 'fusion' )
# Carregar as telas.
TelaPrincipal = uic.loadUi("views\\TelaPrincipalDoSistema.ui")
aviso = uic.loadUi("views\\avisosnovos.ui")
tela_progresso = uic.loadUi("views\\barradeprogreço.ui")
telaDeLogin = uic.loadUi("views\\teladelogin.ui")
telaDeEmail = uic.loadUi("views\\telaDeEmail2.ui")
telaDeVendas = uic.loadUi("views\\teladevendas.ui")
tela_cadastro = uic.loadUi("Views\\tela_cadastro.ui")

TelaPrincipal.consultar_cnpj.clicked.connect(consultarcnpj)# Conectando o click dos botões nas funções
TelaPrincipal.verificaCep.clicked.connect(virificacep)
TelaPrincipal.pushButton_7.clicked.connect(pesquisarProduto)
TelaPrincipal.salvarCliente.clicked.connect(cadcliente)
TelaPrincipal.pushButton_15.clicked.connect(vendasAvista)
TelaPrincipal.geraexcel.clicked.connect(geraRelatorioVendasEntSaida)
TelaPrincipal.enviaemail.clicked.connect(enviaemail)
TelaPrincipal.realizarVendas.clicked.connect(realizarvendas)
tela_progresso.pushButton.clicked.connect(fecharbarradeprogreco)
aviso.sairDaTelaAteno.clicked.connect(fecharteladeaviso)
telaDeLogin.pushButton.clicked.connect(chama_segunda_tela)
TelaPrincipal.realizarVendas_2.clicked.connect(cadastrar_produtos)
TelaPrincipal.pushButton_10.clicked.connect(vendas_parceladas)
TelaPrincipal.deletarProduto.clicked.connect(deletarProduto)
telaDeLogin.RecuperarSenha.clicked.connect(recuperasenhalogin)
telaDeLogin.pushButton_2.clicked.connect(tela_cadastrousuario)
tela_cadastro.pushButton.clicked.connect(cadastrar_usuario)
TelaPrincipal.cadEmpresa.clicked.connect(cadastrar_empresa)
telaDeEmail.pushButton_3.clicked.connect(arquivoaserenviado)
telaDeEmail.enviaremail.clicked.connect(enviaremailcomarquivo)
TelaPrincipal.geraexcel_2.clicked.connect(gerarrelatorioprodutos)
telaDeVendas.finalizar_2.clicked.connect(vender_produto)

descricaoDopagamento = pd.read_sql('select descricao from tipo_pagamento', conexao)# Pegando o Tipo de Pafamento 
TelaPrincipal.comboBox_6.addItems(descricaoDopagamento['descricao'])

descricaoDaEntrsaida = pd.read_sql('SELECT descricao FROM etrada_saida;', conexao)# Pegando o tipo de pagamento
TelaPrincipal.comboBox_8.addItems(descricaoDaEntrsaida['descricao'])

descreicaoDaCategotia = pd.read_sql('select descricao from categorias', conexao)# Pegando o as categorias
TelaPrincipal.comboBox_7.addItems(descreicaoDaCategotia['descricao'])

telaDeLogin.lineEdit_2.setEchoMode(QtWidgets.QLineEdit.Password)# Inplementando campo senha
telaDeLogin.show()
app.exec()