from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QMessageBox
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



try:
    global conexao
    with open('Config\\config.json') as f:# Abrindo o arquivo que contem a string de conexao com o DB
        entrada = json.load(f)# lendo o json e armazenando em uma variavel
    
    

    conexao = mysql.connector.Connect(# string de conexao
        host=entrada["host"],
        user=entrada["user"],
        password=entrada["password"],
        database=entrada["database"],
        auth_plugin=entrada["auth_plugin"]
    )
    with open('logs\\Sistema_De_Vendas_login.txt', 'w') as arquivo:
        arquivo.write('Sistema_De_Vendas: Aplicação conectando ao banco de dados\n')
        arquivo.write('Sistema_De_Vendas: Aplicação conectou ao banco de dados\n')
except Exception as erro:
    with open('logs\\Sistema_De_Vendas_erro.txt', 'w') as arquivo:
        arquivo.write('Sistema_De_Vendas: "Erro: \n"{}\n'.format(erro))# Abrindo o arquivo que contem a string de conexao com o DB
    


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

        with open('logs\\Sistema_De_Vendas_login.txt', 'w') as arquivo:
            arquivo.write('Sistema_De_Vendas: Aplicação conectando ao banco de dados\n')
            arquivo.write('Sistema_De_Vendas: Aplicação conectou ao banco de dados\n')
            arquivo.write('Sistema_De_Vendas: Aplicação Entrou\n')

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
        with open('logs\\Sistema_De_Vendas_catalogaprodutos.txt', 'w') as arquivo:
            arquivo.write('Sistema_De_Vendas: Aplicação catalogou todos os produtos\n')
            

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
        
        with open('logs\\Sistema_De_Vendas_vendas.txt', 'w') as arquivo:
            arquivo.write('Sistema_De_Vendas: Aplicação ccatalogou todas as vendas\n')
            
        
        datadehoje = dt.datetime.now()# Estou pegando a data do dia nessa variavél datavenda
        
        TelaPrincipal.dateEdit_5.setDate(datadehoje)# Estou setando o valor da variavél no objeto da data
        TelaPrincipal.dateEdit_4.setDate(datadehoje)
        TelaPrincipal.dateEdit.setDate(datadehoje)
        TelaPrincipal.dateEdit.setDate(datadehoje)
        
        cursor.close()
        
    except Exception as indexx:
        QMessageBox.warning(aviso, "Aviso", "{}".format(indexx))
        with open('logs\\Sistema_De_Vendas_loginErro.txt', 'w') as arquivo:
            arquivo.write('Sistema_De_Vendas: IndexErro {}\n'.format(indexx))
        return
    


def virificacep():

    try:
        cep = TelaPrincipal.cepCliente.text()

        if not cep:
            with open('logs\\Sistema_De_Vendas_ConsultarCep.txt', 'w') as arquivo:
                arquivo.write('Cep não informado.')
            return False
        else:
            endereco = get_address_from_cep(cep, webservice=WebService.APICEP)
        with open('logs\\Sistema_De_Vendas_ConsultarCep.txt', 'w') as arquivo:
            arquivo.write('{}\n'.format(endereco))
    
    except exceptions.CEPNotFound as notfound:
        QMessageBox.warning(aviso,"Aviso", "{}".format(notfound))
        return

    except exceptions.ConnectionError as conctcaoerro:
        QMessageBox.warning(aviso, "Aviso", "{}".format(conctcaoerro))
        return

    except exceptions.Timeout as tempo:
        QMessageBox.warning(aviso, "Aviso", "{}".format(tempo))
        return

    except exceptions.HTTPError as erro:
        QMessageBox.warning(aviso, "Aviso", "{}".format(erro))
        return

    except exceptions.BaseException as base:
        QMessageBox.warning(aviso, "Aviso", "{}".format(base))
        return

    except ValueError as erru:
        QMessageBox.warning(aviso, "Aviso", "{}".format(erru))
        return
    except Exception as erro:
        QMessageBox.warning(aviso, "Aviso", "{}".format(erro))
        

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
        if verificaCpfExiste.empty ==True:

        
            cursor = conexao.cursor()

            df = pd.read_sql(f"SELECT idestado from estados where sigla = '{estadodocliente}'", conexao)
            if df.empty == True:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Informe o estado.\n\n")
                return
            else:
                estadodocliente = (df['idestado'][0])
            

            df = pd.read_sql(f"SELECT idcidade FROM cidades WHERE nome = '{cidadedocliente}'", conexao)
            
            if df.empty == True:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Informe a cidade.\n\n")
                return
            else:
                cidadedocliente = (df['idcidade'][0])
            

            if not bairrodocliente:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Informe o bairro\n\n")
                return

            pegaridbairro = pd.read_sql(f"SELECT idbairro FROM bairros WHERE nome = '{bairrodocliente}'", conexao)
            if pegaridbairro.empty ==True:
                
                cursor.execute(f"""INSERT INTO bairros (cidades_estados_idestado, cidades_idcidade, nome) VALUES({estadodocliente}, {cidadedocliente}, '{bairrodocliente}');""")
                conexao.commit()
                return
            else:
                pegaridbairro = (pegaridbairro['idbairro'][0])

            if not sitedocliente:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Caso o cliente não possua um site\n preencha o campo com NÃO INFORMADO")
                return

            if not emaildocliente:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Caso o cliente não possua um email\n preencha o campo com NÃO INFORMADO")
                return

            if not nomedocliente:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Informe o nome do cliente")
                return

            if not numerodocliente:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Informe o numero da residencia.")
                return

            if not celulardocliente:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Informe um numero de Celular.")
                return

            if not cpfdocliente:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Informe o CPF do cliente.")
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
                pegaridbairro,
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

            QMessageBox.warning(TelaPrincipal,"mensagem de alerta", "Cliente cadastrado com sucesso.")
            
        else: 
            verificaCpfExiste = (verificaCpfExiste['cpf_cnpj'][0])
            if verificaCpfExiste == cpfdocliente:
                QMessageBox.warning(TelaPrincipal, "Aviso", "Cliente ja cadastrado")
            return
    except Exception as e:
        QMessageBox.warning(aviso,"Aviso", "{}".format(e))
        return




def geraRelatorioVendasEntSaida():
    try:

        datavenda1 = TelaPrincipal.dateEdit_4.text()
        datavenda1 = dt.datetime.strptime(datavenda1, '%d/%m/%Y')
        datavenda1 = datavenda1.strftime('%Y-%m-%d')

        datavenda2 = TelaPrincipal.dateEdit_5.text()
        datavenda2 = dt.datetime.strptime(datavenda2, '%d/%m/%Y')
        datavenda2 = datavenda2.strftime('%Y-%m-%d')
        
        ent_sai = TelaPrincipal.comboBox_8.currentText()
        if ent_sai == 'Entrada':
            ent_sai = int(1)
        elif ent_sai == 'Saida':
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
        QMessageBox.warning(aviso,"Aviso", "{}".format(er))
        
    except Exception as erro:
        QMessageBox.information(aviso,"Aviso", "{}".format(erro))

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
        QMessageBox.information(aviso,"Informação", "{}".format(er))
        
    except Exception as eerro:
        QMessageBox.information(aviso,"Informação", "{}".format(eerro))

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
    except Exception as erro:
        QMessageBox.warning(aviso,"Aviso", "{}".format(erro))
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


def vender_produto():

    try:
        pass
    except Exception as e:
        QMessageBox.warning(aviso, 'Aviso','{}'.format(e))
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


        if not estoque:
            QMessageBox.information(aviso, 'Aviso','Preencha o campo ESTOQUE.')
            return
        cursor = conexao.cursor()
        SQL_produtos = """INSERT INTO produtos (categorias_idcategoria, descricao, preco, observacao,marca,referencia, dt_entrada)
         VALUES  ('{}', '{}', {}, '{}','{}', '{}', '{}')""".format(categoria, descricao, preco, observacao, marca, ref, datadaentrada)
        if not preco:
            QMessageBox.information(aviso, 'Aviso','Preencha os campos vazios EX: Preço')
            return
        elif not descricao:
            QMessageBox.information(aviso, 'Aviso','Capo DESCRIÇÃO é obrigatorio.')
            return
        cursor.execute(SQL_produtos)# Executando o sql para cadastrar o novo produto
        conexao.commit()
 
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
    except Exception as erro:
        QMessageBox.warning(aviso, 'Aviso','{}'.format(erro))
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

    except Exception as erro:
        QMessageBox.warning(aviso, 'Aviso','{}'.format(erro))
        return


def deletarProduto():
    try:
        codigoestoque = TelaPrincipal.codEstoque.text()
        codigoDoproduto = TelaPrincipal.codProduto_2.text()
        if not codigoestoque:
            QMessageBox.information(aviso, 'Aviso','Preencha o campo Codigo/Estoque')
            return
        elif not codigoDoproduto:
            QMessageBox.information(aviso, 'Aviso','Preencha o campo Codigo/Produto')

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

        QMessageBox.information(aviso, 'Aviso','Produto deletado com exito {}'.format(codigoDoproduto))
    except Exception as erro:
        if not codigoestoque:
            QMessageBox.information(aviso, 'Aviso','Preencha o campo Codigo/Estoque')
        elif not codigoDoproduto:
            QMessageBox.information(aviso, 'Aviso','Preencha o campo Codigo/Produto')
        else:
            QMessageBox.warning(aviso, 'Aviso','{}'.format(erro))
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
        QMessageBox.information(aviso, 'Aviso','{}'.format(erro))
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
            QMessageBox.information(aviso, 'Aviso','Email não localizado')
            return
        elif email == 'None':
            QMessageBox.information(aviso, 'Aviso','Email não localizado')
            return
        else:
            email.Send()
            QMessageBox.information(aviso, 'Aviso','Sua senha foi enviada para o email \n{}'.format(email2))
            
                
    except Exception as erro:
        QMessageBox.warning(aviso, 'Atenção', '{}'.format(erro))
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
    except Exception as erros:
        QMessageBox.warning(aviso, 'Atenção', '{}'.format(Erros))
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

        VerificaEmpresa = pd.read_sql(f"select cnpj from empresa where cnpj = '{cnpj}'", conexao)
        if VerificaEmpresa.empty == True:

            idcidade = pd.read_sql(f"select idcidade from cidades where nome ='{idcidade}'", conexao)
            if idcidade.empty == True:
                QMessageBox.warning(aviso, 'Atenção', 'Cidade não informada')
                return
            else:
                idcidade = (idcidade['idcidade'][0])

            idestado = pd.read_sql(f"select idestado from estados where sigla ='{idestado}'", conexao)
            if idestado.empty == True:
                QMessageBox.warning(aviso, 'Atenção', 'Estado não informado')
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

            
            cursor = conexao.cursor()
            cursor.execute("""INSERT INTO think.empresa (cnpj, razao_social, nome_fantazia, tipo_empresa, atividade_principal, natureza_juridica, atividade_secundaria, situacao, capital_social, cep, complemento, email_empresa, telefone, abertura_empresa, porte_empresa, idbairro, idcidade, idestado)
                            VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', {}, '{}', '{}', '{}', '{}', '{}', '{}', {}, {}, {})""".format(cnpj, razao_social, nome_fantazia, tipo_empresa, atividade_principal, natureza_juridica, atividade_secundaria, situacao, capital_socia, cep_empresa, complemento, email_empresa, telefone_empresa, abertura_empresa, porte_empresa, idbairro, idcidade, idestado))
            conexao.commit()
            cursor.close()
            QMessageBox.information(aviso, 'Informação', 'Empresa cadastrada {}'.format(razao_social))
        
        else:
            VerificaEmpresa = (VerificaEmpresa['cnpj'][0])
            if VerificaEmpresa == cnpj:
                QMessageBox.warning(aviso, 'Informação', 'Empresa já cadastrada anteriormente\n{}'.format(razao_social))
                qInfo("{}".format(VerificaEmpresa))

                return 
        return
    except Exception as erro:
        qDebug('{}'.format(erro))
        QMessageBox.warning(aviso, 'Informação', '{}'.format(erro))
   

def arquivoaserenviado():
   
    try:
        salvar = QtWidgets.QFileDialog.getOpenFileName()[0]# Pegando path co arquivo a ser enviado
        telaDeEmail.lineEdit.setText(salvar)# Setando o caminho em um lineEdit
    except Exception as erro:
        QMessageBox.information(aviso, 'Informação', '{}'.format(erro))

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
        
        QMessageBox.information(TelaPrincipal,"Aviso", "Email enviado com sucesso para {}".format(emailDestinatario))
    except Exception as erro:
        QMessageBox.warning(aviso,"Aviso", "{}".format(erro))
        return

def atualizarcliente():
    clientes.show()

def atualizarclientenodb():
    QMessageBox.critical(aviso,"Aviso", "Essa função ainda não esta em funcionamento.")
    return
def deletarregistro():
    QMessageBox.critical(aviso,"Aviso", "Essa função ainda não esta em funcionamento.")
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
clientes = uic.loadUi("Views\\Clientes.ui")

TelaPrincipal.consultar_cnpj.clicked.connect(consultarcnpj)# Conectando o click dos botões nas funções
TelaPrincipal.verificaCep.clicked.connect(virificacep)
TelaPrincipal.pushButton_7.clicked.connect(pesquisarProduto)
TelaPrincipal.salvarCliente.clicked.connect(cadcliente)
TelaPrincipal.pushButton_15.clicked.connect(vendasAvista)
TelaPrincipal.geraexcel.clicked.connect(geraRelatorioVendasEntSaida)
TelaPrincipal.enviaemail.clicked.connect(enviaemail)
TelaPrincipal.realizarVendas.clicked.connect(realizarvendas)
tela_progresso.pushButton.clicked.connect(fecharbarradeprogreco)
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
TelaPrincipal.salvarCliente_3.clicked.connect(atualizarcliente)
clientes.atualizarcliente.clicked.connect(atualizarclientenodb)
clientes.deletarregistro.clicked.connect(deletarregistro)

descricaoDopagamento = pd.read_sql('select descricao from tipo_pagamento', conexao)# Pegando o Tipo de Pafamento 
TelaPrincipal.comboBox_6.addItems(descricaoDopagamento['descricao'])

descricaoDaEntrsaida = pd.read_sql('SELECT descricao FROM etrada_saida;', conexao)# Pegando o tipo de pagamento
TelaPrincipal.comboBox_8.addItems(descricaoDaEntrsaida['descricao'])

descreicaoDaCategotia = pd.read_sql('select descricao from categorias', conexao)# Pegando o as categorias
TelaPrincipal.comboBox_7.addItems(descreicaoDaCategotia['descricao'])

telaDeLogin.lineEdit_2.setEchoMode(QtWidgets.QLineEdit.Password)# Inplementando campo senha
telaDeLogin.show()
app.exec()