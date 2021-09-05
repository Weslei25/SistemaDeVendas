from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import * 
from PyQt5.QtGui import * 
from PyQt5.QtCore import *
from PyQt5.QtWidgets import QMessageBox
import mysql.connector
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pycep_correios import get_address_from_cep, WebService, exceptions
import requests
import datetime as dt
import pandas as pd
import json
import bcrypt
import logging
from PyQt5.QtWidgets import * 
from PyQt5.QtCore import Qt, QSortFilterProxyModel
from PyQt5.QtGui import QStandardItem, QStandardItemModel
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




def chama_segunda_tela():# função responsavel por chamar a tela principal
    
    telaDeLogin.label_4.setText("")# Sempre limpo o campo de avisos para que quando o usuario corrigir o erro o campo esteja limpo
    nome_usuario = telaDeLogin.lineEdit.text()# Aqui pego o nome de usuario
    senha = telaDeLogin.lineEdit_2.text()# Aqui pego a senha
    

    try:

        if not nome_usuario:
            telaDeLogin.label_4.setText("Usuario ou Senha incorretos.")
            return
        elif not senha:
            telaDeLogin.label_4.setText("Usuario ou Senha incorretos.")
            return

        verificaUsuario = pd.read_sql(f"select nome from usuarios where nome='{nome_usuario}'", conexao)

        if verificaUsuario.empty == True:
            telaDeLogin.label_4.setText("Usuario não encontrado.")
            return
        else:
            verificaUsuario = (verificaUsuario['nome'][0])

        cursor = conexao.cursor()
        cursor.execute("select senha FROM usuarios WHERE nome ='{}'".format(nome_usuario))
        pegasenha = cursor.fetchall()
        verificaSenha = pegasenha[0][0]
        
        verificaSenha = (verificaSenha).encode('utf-8')
    
        senha2 = (senha).encode('utf-8')

       
        if bcrypt.hashpw(senha2, verificaSenha)==verificaSenha:
            telaDeLogin.close()
            TelaPrincipal.show()
        else:
            telaDeLogin.label_4.setText("Usuario ou Senha incorretos.")
            return
        TelaPrincipal.usuarios.setText('Usuario Logado: {}'.format(nome_usuario))
        datahoje = dt.datetime.now()

        nomedousuario = pd.read_sql(f"select idusuarios from usuarios where nome='{nome_usuario}'", conexao)
        nomedousuario = (nomedousuario['idusuarios'][0])
        cursor = conexao.cursor()
        cursor.execute(f"INSERT INTO log_usuario(descricao, idusuarios, dt_logusuario) VALUES ('NULL', {nomedousuario}, '{datahoje}');")
        conexao.commit()

        logging.info("Logando na aplicação com o usuario {}".format(nome_usuario))
        logging.info("Aplicação entrou ")
        

        def catalogarProdutos():
            logging.info("Aplicação tenta catalogar produtos")
            try:
                cursor = conexao.cursor()
                sql_cataloga = """
                SELECT idproduto, descricao, (DATE_FORMAT(dt_entrada , '%d/%m/%Y')), format(preco,2,'de_DE'), observacao, codbarras, marca,
                referencia from produtos order by idproduto DESC limit 1200"""
                
                cursor.execute(sql_cataloga)
                with open('sql\\SQL_catalogarProdutos.sql', 'w') as arquivo:
                    arquivo.write(f'{sql_cataloga}')

                dados_lidos1 = cursor.fetchall()
                TelaPrincipal.tableWidget.setRowCount(len(dados_lidos1))
                TelaPrincipal.tableWidget.setColumnCount(8)
                for i in range(0, len(dados_lidos1)):
                    for j in range(0, 8):
                        TelaPrincipal.tableWidget.setItem(
                            i, j, QtWidgets.QTableWidgetItem(str(dados_lidos1[i][j])))
                logging.info("Todos os produtos catalogados")
            except Exception as erro:
                logging.exception(erro)

        catalogarProdutos()
        
        logging.info("Aplicação tenta catalogar vendas")
        idnegociacao = TelaPrincipal.tiponegociacao.currentText()
        idvendedor = TelaPrincipal.idvendedor.currentText()
        id_entrada_saida = 1
        idpagamento = 15
        idnegociacao = pd.read_sql(f"select idtipo_negociacao from tipo_negociacao where descricao='{idnegociacao}'", conexao)
        idnegociacao = (idnegociacao['idtipo_negociacao'][0])

        idvendedor = pd.read_sql(f"select idvendedor from vendedores where nome ='{idvendedor}'", conexao)
        idvendedor = (idvendedor['idvendedor'][0])

        cursor = conexao.cursor()
        cursor.execute(f"""select idvenda, nomecliente,
            (select descricao from tipo_negociacao where idtipo_negociacao={idnegociacao}),
            (select nome from vendedores where idvendedor={idvendedor}), (DATE_FORMAT(data_venda , '%d/%m/%Y')), (DATE_FORMAT(dat_venv_fatuura , '%d/%m/%Y')), nomeproduto,
            quantproduto, format(precoproduto,2,'de_DE'), format(descproduto,2,'de_DE'), format(vlr_total,2,'de_DE'), (select descricao from tipo_pagamento where idpagamento={idpagamento}),
            vezesdeparcelas, observacao, (select descricao from etrada_saida where id_entrada_saida={id_entrada_saida}) FROM vendas order by idvenda DESC limit 1000000;""")


        sql_vendas1 = cursor.fetchall()

        TelaPrincipal.tableWidget_5.setRowCount(len(sql_vendas1))
        TelaPrincipal.tableWidget_5.setColumnCount(15)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 15):
                TelaPrincipal.tableWidget_5.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))
        
        logging.info("Aplicação carregou todas as vendas")
            
        
        datadehoje = dt.datetime.now()# Estou pegando a data do dia nessa variavél datavenda
        
        TelaPrincipal.dateEdit_5.setDate(datadehoje)# Estou setando o valor da variavél no objeto da data
        TelaPrincipal.dateEdit_4.setDate(datadehoje)
        TelaPrincipal.dateEdit.setDate(datadehoje)
        TelaPrincipal.dateEdit.setDate(datadehoje)
        
        cursor.close()
        
    except Exception as erro:
        QMessageBox.warning(aviso, "Aviso", "{}".format(erro))
        logging.exception(erro)
        return
       


def virificacep():

    try:
        cep = TelaPrincipal.cepCliente.text()
        logging.info("Aplicação tenta buscar o cep {}".format(cep))
        if not cep:
            logging.info("Cep not found")
            return False
        else:
            endereco = get_address_from_cep(cep, webservice=WebService.APICEP)
        logging.info("{}".format(endereco))
    
    except exceptions.CEPNotFound as notfound:
        logging.exception(notfound)
        QMessageBox.warning(aviso,"Aviso", "{}".format(notfound))
        return

    except exceptions.ConnectionError as conctcaoerro:
        logging.exception(conctcaoerro)
        QMessageBox.warning(aviso, "Aviso", "{}".format(conctcaoerro))
        return

    except exceptions.Timeout as tempo:
        logging.exception(tempo)
        QMessageBox.warning(aviso, "Aviso", "{}".format(tempo))
        return

    except exceptions.HTTPError as erro:
        logging.exception(erro)
        QMessageBox.warning(aviso, "Aviso", "{}".format(erro))
        return

    except exceptions.BaseException as base:
        logging.exception(base)
        QMessageBox.warning(aviso, "Aviso", "{}".format(base))
        return

    except ValueError as erru:
        logging.exception(erru)
        QMessageBox.warning(aviso, "Aviso", "{}".format(erru))
        return
    except Exception as erros:
        logging.exception(erros)
        QMessageBox.warning(aviso, "Aviso", "{}".format(erros))
        

    TelaPrincipal.enderecoCliente.setText(endereco['logradouro'])
    TelaPrincipal.bairroCliente.setText(endereco['bairro'])
    TelaPrincipal.cidadeCliente.setText(endereco['cidade'])
    TelaPrincipal.compleCliente.setText(endereco['logradouro'])
    TelaPrincipal.estadoDocliente.setText(endereco['uf'])
    logging.info("Cep buscado e encontrado")

def cadcliente():
    try:
        logging.info("Entra função de cadastro de clientes")
        # Variaveis para cadastro de clientes
        nomedocliente = str(TelaPrincipal.nomeDoCliente.text())
        cepdocliente = str(TelaPrincipal.cepCliente.text())
        cidadedocliente = str(TelaPrincipal.cidadeCliente.text())
        bairrodocliente = str(TelaPrincipal.bairroCliente.text())
        ruadocliente = str(TelaPrincipal.enderecoCliente.text())
        numerodocliente = str(TelaPrincipal.numeroCliente.text())
        complemento = str(TelaPrincipal.compleCliente.text())
        estadodocliente = str(TelaPrincipal.estadoDocliente.text())
        celulardocliente = str(TelaPrincipal.telCell.text())
        telefoneresidencial = str(TelaPrincipal.telResid.text())
        categoriadocliente = str(TelaPrincipal.catCliente.currentText())
        cpfdocliente = str(TelaPrincipal.cpfCliente.text())
        rgdocliente = str(TelaPrincipal.rgCliente.text())
        sitedocliente = str(TelaPrincipal.siteCliente.text())
        emaildocliente = str(TelaPrincipal.emailDoCliente.text())
        informacaocliente = str(TelaPrincipal.infoCliente.toPlainText())

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
                TelaPrincipal.cidadeCliente.setStyleSheet('background-color: rgb(255, 73, 73);')
                return
            else:
                cidadedocliente = (df['idcidade'][0])
            

            if not bairrodocliente:
                QMessageBox.warning(TelaPrincipal, "Atenção", "Informe o bairro\n\n")
                return

            pegaridbairro = pd.read_sql(f"select idbairro from bairros where nome='{bairrodocliente}' and cidades_idcidade='{cidadedocliente}'", conexao)
            if pegaridbairro.empty ==True:
                logging.info("Bairro não encontrado no banco de dados, sendo feito o cadastro agora")
                cursor.execute(f"""INSERT INTO bairros (cidades_estados_idestado, cidades_idcidade, nome) VALUES({estadodocliente}, {cidadedocliente}, '{bairrodocliente}');""")
                conexao.commit()
                logging.info("Bairro cadastrado com sucesso")
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
            tel_secund, email, site, observacao) values(
                {}, {}, {}, '{}', '{}', 'F', '{}', '{}', '{}', '{}',
                '{}', '{}', '{}', '{}', '{}', '{}', '{}') """.format(
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
                sitedocliente, informacaocliente)

            cursor.execute(sql_cliente)
            conexao.commit()
            cursor.close()

            QMessageBox.information(TelaPrincipal,"mensagem de alerta", "Cliente cadastrado com sucesso.")
            logging.info("Novo cliente cadastrado")
            TelaPrincipal.cpfCliente.setStyleSheet('')
            
        else: 
            verificaCpfExiste = (verificaCpfExiste['cpf_cnpj'][0])
            if verificaCpfExiste == cpfdocliente:
                QMessageBox.information(TelaPrincipal, "Aviso", "Cliente ja cadastrado anteriormente \nCPF {}".format(cpfdocliente))
                logging.warning("Tentou cadastrar cliente ja cadastrado")
                TelaPrincipal.cpfCliente.setStyleSheet('background-color: rgb(255, 73, 73);')
                # TelaPrincipal.cpfCliente.setStyleSheet('color: rgb(255, 255, 255);')
                
            return
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.warning(aviso,"Aviso", "{}".format(erro))
        return




def geraRelatorioVendasEntSaida():
    try:
        logging.info("Chama funcao nome -> geraRelatorioVendasEntSaida()")
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
        logging.info("Relatorio foi criado com sucesso")
    except ValueError as erros:
        logging.exception(erros)
        QMessageBox.warning(aviso,"Aviso", "{}".format(erros))
    
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.warning(aviso,"Aviso", "{}".format(erro))

def gerarrelatorioprodutos():
    
    try:
        logging.info("Entra função -> gerarrelatorioprodutos")
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
    except ValueError as erro:
        logging.exception(erro)
        QMessageBox.information(aviso,"Informação", "{}".format(erro))
        return
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.information(aviso,"Informação", "{}".format(erro))
        return

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
        logging.exception(erro)
        QMessageBox.warning(aviso,"Aviso", "{}".format(erro))
        return

def realizarvendas():
    
    try:

        vendedor = TelaPrincipal.nomevendedor.text()
        entrada = TelaPrincipal.entrada.text()
        saida = TelaPrincipal.saida.text()
        clienteVenda = TelaPrincipal.clienteVendas.text()
        cnpjCpf = TelaPrincipal.cpf_cnpj_2.text()
        prosutoVenda = TelaPrincipal.produto_2.text()
        quantidadeVenda = TelaPrincipal.quantidade_2.text()
        cdDeBarras = TelaPrincipal.codigo_de_barras_2.text()
        descontoVEnda = TelaPrincipal.desconto_2.text()
        PorcentoVenda = TelaPrincipal.porcento_2.text()
        categoriaVenda = TelaPrincipal.categoriasVendas.currentText()
        quantidadeitens = TelaPrincipal.qtItens.text()
        quantidadeProdutos = TelaPrincipal.qtprodutos.text()
        troco = TelaPrincipal.troco.text()
        saldoDevedor = TelaPrincipal.saldoDevedor.text()
        totalDeDesconto = TelaPrincipal.descontoTotal.text()
        TotalVendas = TelaPrincipal.total.text()

        
        QMessageBox.information(aviso, "Ok", "Ok Cambada do caralho")
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.warning(aviso, 'Aviso','{}'.format(erro))


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
        codBarras = str(TelaPrincipal.codBarras.text())


        pegarcodbarras = pd.read_sql(f"select codbarras from produtos where codbarras ='{codBarras}'", conexao)
        if pegarcodbarras.empty == True:

            pegarcategoria = pd.read_sql(f"select idcategoria from categorias c where descricao ='{categoria}'", conexao)
            categoria = (pegarcategoria['idcategoria'][0])
            
            if not estoque:
                QMessageBox.information(aviso, 'Aviso','Preencha os campos obrigatorios')
                TelaPrincipal.estoque.setStyleSheet('background-color: rgb(255, 0, 0);color: rgb(255, 255, 255);')
                return
            
            cursor = conexao.cursor()
            SQL_produtos = """INSERT INTO produtos (categorias_idcategoria, descricao, preco, observacao,marca,referencia, dt_entrada, codbarras)
            VALUES  ('{}', '{}', {}, '{}','{}', '{}', '{}','{}')""".format(categoria, descricao, preco, observacao, marca, ref, datadaentrada, codBarras)
            if not preco:
                QMessageBox.information(aviso, 'Aviso','Preencha os campos obrigatorios')
                TelaPrincipal.preco.setStyleSheet('background-color: rgb(255, 0, 0);color: rgb(255, 255, 255);')
                return
            if not descricao:
                QMessageBox.information(aviso, 'Aviso','Preencha os campos obrigatorios')
                TelaPrincipal.descricao.setStyleSheet('background-color: rgb(255, 0, 0);color: rgb(255, 255, 255);')
                return
            if not codBarras:
                QMessageBox.information(aviso, 'Aviso','Preencha os campos obrigatorios')
                TelaPrincipal.codBarras.setStyleSheet('background-color: rgb(255, 0, 0);color: rgb(255, 255, 255);')
                return
            cursor.execute(SQL_produtos)# Executando o sql para cadastrar o novo produto
            conexao.commit()
    
            cursor.execute("SELECT MAX(idproduto) FROM produtos")
            produtoult = cursor.fetchall()
            tratadoproduto = produtoult[0][0]
            cursor.execute("INSERT INTO estoque (estoque, produtos_idproduto) values ({},{})".format(estoque,tratadoproduto))
            conexao.commit()
            
            cursor.execute("""SELECT idproduto, descricao, (DATE_FORMAT(dt_entrada , '%d/%m/%Y')), format(preco,2,'de_DE'), observacao, codbarras, marca,
            referencia from produtos order by idproduto DESC limit 1200""")
            sql_tprodu = cursor.fetchall()

            TelaPrincipal.tableWidget.setRowCount(len(sql_tprodu))
            TelaPrincipal.tableWidget.setColumnCount(8)

            for i in range(0, len(sql_tprodu)):
                for j in range(0, 8):
                    TelaPrincipal.tableWidget.setItem(
                        i, j, QtWidgets.QTableWidgetItem(str(sql_tprodu[i][j])))

            cursor.close()

            """
            TelaPrincipal.lineEdit_2.setText("")
            TelaPrincipal.lineEdit_3.setText("")
            TelaPrincipal.lineEdit_4.setText("")
            TelaPrincipal.lineEdit_5.setText("")
            TelaPrincipal.lineEdit_9.setText("")
            TelaPrincipal.codBarras.setText("")
            """
            
            TelaPrincipal.codBarras.setStyleSheet('background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);')
            TelaPrincipal.estoque.setStyleSheet('background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);')
            TelaPrincipal.preco.setStyleSheet('background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);')
            TelaPrincipal.descricao.setStyleSheet('background-color: rgb(255, 255, 255);color: rgb(0, 0, 0);')
        else:
            descricao = (pegarcodbarras['codbarras'][0])
            logging.warning(f"Produto ja cadastrado anteriormente {descricao}")
            QMessageBox.warning(aviso, 'Aviso','Produto ja cadastrado anteriormente\nCodigo de Barras {}\nCaso não seja o mesmo produto verifique o codigo de barras'.format(descricao))
            TelaPrincipal.codBarras.setStyleSheet('background-color: rgb(255, 0, 0);color: rgb(255, 255, 255);')
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.critical(aviso, 'Aviso','{}'.format(erro))
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
        logging.exception(erro)
        QMessageBox.information(aviso, 'Aviso','{}'.format(erro))
        return


def deletarProduto():
    try:
        codigoestoque = TelaPrincipal.codEstoque.text()
        codigoDoproduto = TelaPrincipal.codProduto_2.text()
        if not codigoestoque:
            QMessageBox.information(aviso, 'Aviso','Para atualizar o estoque de um produto é necessario\n passar o codigo do PRODUTO e o codigo do ESTOQUE referente\n ao produto')
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
        logging.exception(erro)
        if not codigoestoque:
            QMessageBox.information(aviso, 'Aviso','Preencha o campo Codigo/Estoque')
            return
        elif not codigoDoproduto:
            QMessageBox.information(aviso, 'Aviso','Preencha o campo Codigo/Produto')
            return
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
        logging.exception(erro)
        QMessageBox.information(aviso, 'Aviso','{}'.format(erro))
        return
def recuperausuario():

    nomeUsuario = TelaPrincipal.nomeDoUsuario.text()
    emailUsuario = TelaPrincipal.emailDoUsuario.text()
    novaSenha = TelaPrincipal.senhaDoUsuario.text()
    verificaMesmaSenha = TelaPrincipal.c_senhaUsuario.text()
    dataDaAtualizacao = dt.datetime.now()

    if not nomeUsuario:
        TelaPrincipal.avisosUsuario.setText("Preencha todos os campos.")
        return
    if not emailUsuario:
        TelaPrincipal.avisosUsuario.setText("Preencha todos os campos.")
        return
    if not novaSenha:
        TelaPrincipal.avisosUsuario.setText("Preencha todos os campos.")
        return
    if not verificaMesmaSenha:
        TelaPrincipal.avisosUsuario.setText("Preencha todos os campos.")
        return

    if novaSenha != verificaMesmaSenha:
        TelaPrincipal.avisosUsuario.setText("Senhas informadas são diferentes.")
        return
    try:
        hashed = (novaSenha).encode('utf-8')
        hashed = bcrypt.hashpw(hashed, bcrypt.gensalt())
        
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.critical(aviso, 'Atenção', '{}'.format(erro))
    
    try:
        retorno = pd.read_sql(f"select idusuarios from usuarios where nome ='{nomeUsuario}' and email ='{emailUsuario}';", conexao)
        if retorno.empty == True:
            QMessageBox.information(aviso, 'Atenção', 'Usuario não encontrado, verifique se o nome e email estão corretos.\nCaso você não possua um usuario pode criar um na aba de cadastro.')
            return
        else:
            idUsuario = (retorno['idusuarios'][0])
        
        if novaSenha == verificaMesmaSenha:
            cursor = conexao.cursor()
            hashed = (hashed).decode('utf-8')

            cursor.execute(f"update usuarios set senha ='{hashed}', updated_at ='{dataDaAtualizacao}' where idusuarios = {idUsuario};")
            conexao.commit()
        else:
            QMessageBox.warning(aviso, 'Atenção', 'Senhas não são iguais '.format(erro))    
            return
        TelaPrincipal.avisosUsuario.setText("Sua senha foi atualizada com sucesso.")
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.warning(aviso, 'Atenção', '{}'.format(erro))
        return


def cadastrar_usuario():

    
    nome = TelaPrincipal.cadNome.text()
    email = TelaPrincipal.cadEmail.text()
    senha = TelaPrincipal.cadSenha.text()
    c_senha = TelaPrincipal.contraSenha.text()
    data_criacao = dt.datetime.now()

    verificaUsuario = pd.read_sql(f"select nome from usuarios where nome='{nome}'", conexao)
    if verificaUsuario.empty == True:
        pass
    else:
        TelaPrincipal.label_2.setText('Ja existe um usuario com esse nome.')
        return
    if not nome:
        TelaPrincipal.label_2.setText('Preencha todos os campos.')
        return
    if not email:
        TelaPrincipal.label_2.setText('Preencha todos os campos.')
        return
    if not senha:
        TelaPrincipal.label_2.setText('Preencha todos os campos.')
        return
    if not c_senha:
        TelaPrincipal.label_2.setText('Preencha todos os campos.')
        return

    senha = (senha).encode('utf-8')# Codificando a senha para utf-8
    c_senha = (c_senha).encode('utf-8')# Codificando a contra_senha para utf-8
    
    try:
        hashed = bcrypt.hashpw(senha, bcrypt.gensalt())
        
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.critical(aviso, 'Atenção', '{}'.format(erro))

    
    if (senha == c_senha):
        
        try:
            hashed = (hashed).decode('utf-8')
            cursor = conexao.cursor()
            sql_user = """INSERT INTO usuarios (nome, email, senha, created_at)
            VALUES ('{}','{}',"{}", '{}')""".format(nome, email, hashed, data_criacao)
            cursor.execute(sql_user)
            conexao.commit()



            nome = TelaPrincipal.lineEdit.setText("")
            email = TelaPrincipal.lineEdit_2.setText("")
            senha = TelaPrincipal.lineEdit_3.setText("")
            c_senha = TelaPrincipal.lineEdit_4.setText("")

            TelaPrincipal.label_2.setText("Usuario cadastrado com sucesso")
            cursor.close()

        except NameError as erro:
            logging.exception(erro)
            TelaPrincipal.label_2.setText('\n\n{}'.format(erro))
            return
        except IndexError as erro2:
            logging.exception(erro2)
            TelaPrincipal.label_2.setText('\n\n{}'.format(erro2))
            return
        except ValueError as erro3:
            logging.exception(erro3)
            TelaPrincipal.label_2.setText('\n\n{}'.format(erro3))
            return
        except AttributeError as erro4:
           logging.exception(erro4)
           TelaPrincipal.label_2.setText('\n\n{}'.format(erro4))
           return
        except mysql.connector.errors.ProgrammingError as erro5:
            logging.exception(erro5)
            TelaPrincipal.label_2.setText('\n\n{}'.format(erro5))
            return
    # Um else caso somente a senha se estiver errada
    else:
        TelaPrincipal.label_2.setText("As senhas digitadas estão diferentes")
        return
    
    
def pesquisarProduto():
    try:
        pesquisar = TelaPrincipal.lineEdit_12.text()
        cursor = conexao.cursor()
        cursor.execute("""SELECT idproduto, descricao, (DATE_FORMAT(dt_entrada , '%d/%m/%Y')), format(preco,2,'de_DE'), observacao, codbarras, marca,
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
        logging.exception(erros)
        QMessageBox.warning(aviso, 'Atenção', '{}'.format(erros))
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

        nomedacidade = idcidade
        nomedobairro = idbairros
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

            idbairro = pd.read_sql(f"select idbairro from bairros where nome='{nomedobairro}' and cidades_idcidade='{nomedacidade}'", conexao)
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
                TelaPrincipal.cnpj_consulta.setStyleSheet('background-color: rgb(255, 73, 73);')

                return 
        return
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.warning(aviso, 'Informação', '{}'.format(erro))
   

def arquivoaserenviado():
   
    try:
        salvar = QtWidgets.QFileDialog.getOpenFileName()[0]# Pegando path co arquivo a ser enviado
        TelaPrincipal.lineEdit.setText(salvar)# Setando o caminho em um lineEdit
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.information(aviso, 'Informação', '{}'.format(erro))
        return
def enviaremailcomarquivo():
    logging.info("Tenta enviar email")
    
    try:
        # 
        # email =  think_V1@outlook.com
        # senha = weslei080319
        # Email do sistema sistemadevendasecadastro2522@gmail.com

        emailDestinatario = TelaPrincipal.lineEdit_2.text() # pego o destinatario 
        anexodoemail = TelaPrincipal.lineEdit.text()# Recebo o anexo do email se tiver.
        corpoDoEmail = TelaPrincipal.textEdit.toPlainText()# QTextEdit.toPlainText é a propriedade que aceita a quebra de linha no qtextEdit
       
        fromaddr = "sistemadevendasecadastro2522@gmail.com"# Email remetente
        toaddr = emailDestinatario # Email destinatario
        msg = MIMEMultipart()

        msg['From'] = fromaddr 
        msg['To'] = toaddr
        msg['Subject'] = TelaPrincipal.lineEdit_3.text()

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
        logging.info("Email enviado com sucesso")

        QMessageBox.information(TelaPrincipal,"Aviso", "Email enviado com sucesso para \n{}".format(emailDestinatario))
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.warning(aviso,"Aviso", "{}".format(erro))
        return



def atualizarclientenodb():
    QMessageBox.information(aviso,"Aviso", "Essa função ainda esta em faze de desenvolvimento.")
    return
def deletarregistro():
    QMessageBox.information(aviso,"Aviso", "Essa função ainda esta em faze de desenvolvimento.")
    return

def consultas():
    descricaoDopagamento = pd.read_sql('select descricao from tipo_pagamento', conexao)# Pegando o Tipo de Pafamento 
    TelaPrincipal.comboBox_6.addItems(descricaoDopagamento['descricao'])

    descricaoDaEntrsaida = pd.read_sql('SELECT descricao FROM etrada_saida;', conexao)# Pegando o tipo de pagamento
    # TelaPrincipal.comboBox_8.addItems(descricaoDaEntrsaida['descricao'])

    descreicaoDaCategotia = pd.read_sql('select descricao from categorias', conexao)# Pegando o as categorias
    TelaPrincipal.categotiaproduto.addItems(descreicaoDaCategotia['descricao'])

    descreicaoTipoNegociacao = pd.read_sql('select descricao from tipo_negociacao', conexao)# Pegando o as categorias
    TelaPrincipal.tiponegociacao.addItems(descreicaoTipoNegociacao['descricao'])

    descreicaoVendedores = pd.read_sql('select nome from vendedores', conexao)# Pegando o as categorias
    TelaPrincipal.idvendedor.addItems(descreicaoVendedores['nome'])

    descreicaoDaCategotia = pd.read_sql('select tipo_cliente from cat_cliente', conexao)# Adicionando as categorias dos clientes
    # TelaPrincipal.catCliente.addItems(descreicaoDaCategotia['tipo_cliente'])

def acessarindiceclientes():
    TelaPrincipal.stackedWidget.setCurrentIndex(0)
    

def cessarindeceinicial():
    TelaPrincipal.stackedWidget.setCurrentIndex(1)
    

def acessarindicevendas():
    TelaPrincipal.stackedWidget.setCurrentIndex(3)
    

def acessarindiceempresas():
    TelaPrincipal.stackedWidget.setCurrentIndex(4)
    

def acessarindicerelatorios():
    TelaPrincipal.stackedWidget.setCurrentIndex(4)
    

def acessarindiceprodutos():
    TelaPrincipal.stackedWidget.setCurrentIndex(2)
    

def acessarindiceemail():
    TelaPrincipal.stackedWidget.setCurrentIndex(5)
    

def acessarindicevender():
    TelaPrincipal.stackedWidget.setCurrentIndex(6)
    pegaridusuario = pd.read_sql("select idusuarios from log_usuario where idusuarios is not null order by id_logusuario DESC limit 1", conexao)
    pegaridusuario = (pegaridusuario['idusuarios'][0])

    df = pd.read_sql(f"select nome from usuarios where idusuarios='{pegaridusuario}'", conexao)
    pegaridusuario = str(df['nome'][0])

    TelaPrincipal.nomevendedor.setText(pegaridusuario)

def acessarindicereverlogin():
    user = TelaPrincipal.usuarios.text()
    if user == "Usuario Logado: admin":

        TelaPrincipal.stackedWidget.setCurrentIndex(7)
    else:
        QMessageBox.information(aviso, "Aviso", f"{user}, sem permissão para acessar essa configuação\nContate um adimistrador para realizar essa transação")   
        return
def acessarindiceclientecadastrados():
    TelaPrincipal.stackedWidget.setCurrentIndex(9)


def acessarsair():
    exit()


def acessarindicepesquisarprodutos():
    TelaPrincipal.stackedWidget.setCurrentIndex(8)

def tentaracesar():
    carros = pd.read_sql("select descricao from produtos", conexao)
    carros = (carros['descricao'])

    # carros = ('Gol', 'Celta', 'Corsa', 'Uno', 'Fox', 'Cruze', 'Brasilia', 'Saveiro', 'Fusca', 'Hilux', 'Onix')
    modelo = QStandardItemModel(len(carros),1)
    modelo.setHorizontalHeaderLabels(['Carros'])

    for linha, carro in enumerate(carros):    # [(1, 'Gol'), (2,'Celta') ]     
        elemento = QStandardItem(carro)
        modelo.setItem(linha, 0, elemento)

    # global filtro
    filtro = QSortFilterProxyModel()
    filtro.setSourceModel(modelo)
    filtro.setFilterKeyColumn(0)
    filtro.setFilterCaseSensitivity(Qt.CaseInsensitive)

    TelaPrincipal.tableView.setModel(filtro)
    TelaPrincipal.tableView.horizontalHeader().setStyleSheet("font-size: 35px;color: rgb(50, 50, 255);")
    TelaPrincipal.pesquisar.textChanged.connect(filtro.setFilterRegExp)

def consultarempresas(cnpj):

    try:
        cnpj = TelaPrincipal.cnpj_consulta.text()
        identificadorempresa = pd.read_sql(f"select idbairro, idcidade,idestado from empresa where cnpj='{cnpj}'", conexao)
        idbairro = (identificadorempresa['idbairro'][0])
        idcidade = (identificadorempresa['idcidade'][0])
        idestado = (identificadorempresa['idestado'][0])
        cursor = conexao.cursor()
        cursor.execute(f"""select
                    id_empresa, cnpj, razao_social,
                    nome_fantazia, tipo_empresa,
                    atividade_principal, natureza_juridica,
                    atividade_secundaria, situacao, capital_social,
                    cep, complemento, email_empresa, telefone,
                    (DATE_FORMAT(abertura_empresa , '%d/%m/%Y')), porte_empresa,
                    (select nome from bairros where idbairro={idbairro}),
                    (select nome from cidades where idcidade={idcidade}),
                    (select nome from estados where idestado={idestado})
                    from
                    empresa
                    where cnpj='{cnpj}';""")
        dadosrecebidos = cursor.fetchall()

        TelaPrincipal.dadosEmpresaCatalogados.setRowCount(len(dadosrecebidos))
        TelaPrincipal.dadosEmpresaCatalogados.setColumnCount(19)
        

        for i in range(0, len(dadosrecebidos)):
            for j in range(0, 19):
                TelaPrincipal.dadosEmpresaCatalogados.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(dadosrecebidos[i][j])))

    except IndexError as erroindex:
        logging.exception(erroindex)
        QMessageBox.warning(aviso, 'Atenção', 'CNPJ é invalido ou não existe na base de dados')
        return
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.warning(aviso, 'Atenção', '{}'.format(erro))
        return





if __name__ == "__main__":
    
    app = QtWidgets.QApplication([]) # Criando um QApplication que faz a construcao da minha aplicacao.
    app.setStyle ( 'fusion' )
    # Carregar as telas.
    TelaPrincipal = uic.loadUi("Views/TelaDefinitiva.ui")
    aviso = uic.loadUi("Views/avisosnovos.ui")
    tela_progresso = uic.loadUi("views/barradeprogreço.ui")
    telaDeLogin = uic.loadUi("Views/teladelogin.ui")
    # Configuracao de logs
    log_format = '%(asctime)s:%(levelname)s:%(filename)s:%(message)s'
    logging.basicConfig(filename='SistemaDeVendas.log',
                    filemode='w',
                    level=logging.DEBUG,
                    format=log_format,
                    encoding='UTF-8')
    logging.debug("Criando a aplicação")

    # telaDeVendas = uic.loadUi("Views/teladevendas.ui")
    # clientes = uic.loadUi("Views/Clientes.ui")
    TelaPrincipal.finalizarVendas.clicked.connect(realizarvendas)
    TelaPrincipal.consultar_cnpj.clicked.connect(consultarcnpj)# Conectando o click dos botões nas funções
    TelaPrincipal.verificaCep_3.clicked.connect(virificacep)
    TelaPrincipal.pushButton_7.clicked.connect(pesquisarProduto)
    TelaPrincipal.salvarCliente.clicked.connect(cadcliente)
    # TelaPrincipal.pushButton_15.clicked.connect(vendasAvista)
    # TelaPrincipal.geraexcel.clicked.connect(acessarindicerelatorios)
    # TelaPrincipal.enviaemail.clicked.connect(cessarindeceemail)
    # TelaPrincipal.realizarVendas.clicked.connect(realizarvendas)
    tela_progresso.pushButton.clicked.connect(fecharbarradeprogreco)
    telaDeLogin.pushButton.clicked.connect(chama_segunda_tela)
    TelaPrincipal.cadastrarProduto.clicked.connect(cadastrar_produtos)
    # TelaPrincipal.pushButton_10.clicked.connect(vendas_parceladas)
    TelaPrincipal.deletarProduto.clicked.connect(deletarProduto)
    TelaPrincipal.pushButton.clicked.connect(cadastrar_usuario)
    TelaPrincipal.cadEmpresa.clicked.connect(cadastrar_empresa)
    TelaPrincipal.pushButton_3.clicked.connect(arquivoaserenviado)
    TelaPrincipal.enviaremail.clicked.connect(enviaremailcomarquivo)
    TelaPrincipal.geraexcel_2.clicked.connect(gerarrelatorioprodutos)
    # telaDeVendas.finalizar_2.clicked.connect(vender_produto)
    # TelaPrincipal.salvarCliente_3.clicked.connect(acessarindiceclientes)
    TelaPrincipal.atualizarcliente.clicked.connect(atualizarclientenodb)
    TelaPrincipal.deletarregistro.clicked.connect(deletarregistro)
    TelaPrincipal.recuperarSenha.clicked.connect(recuperausuario)
    # TelaPrincipal.acessarProdutos.clicked.connect(acessarindiceprodutos)
    # TelaPrincipal.pushButton_14.clicked.connect(acessarindice)
    TelaPrincipal.actionProdutos.triggered.connect(acessarindiceprodutos)
    TelaPrincipal.actionVendas_2.triggered.connect(acessarindicevendas)
    TelaPrincipal.actionCadastrar_clientes.triggered.connect(acessarindiceclientes)
    TelaPrincipal.actionCadastrar_nova_empresa.triggered.connect(acessarindiceempresas)
    TelaPrincipal.actionPag_inicial.triggered.connect(cessarindeceinicial)
    TelaPrincipal.actionEmail_com_relatorio.triggered.connect(acessarindiceemail)
    TelaPrincipal.actionVender.triggered.connect(acessarindicevender)
    TelaPrincipal.actionCriar_usuario_novo.triggered.connect(acessarindicereverlogin)
    TelaPrincipal.actionRecuperar_acesso.triggered.connect(acessarindicereverlogin)
    TelaPrincipal.actionSair.triggered.connect(acessarsair)
    TelaPrincipal.actionClientes_cadastrados.triggered.connect(acessarindiceclientecadastrados)
    TelaPrincipal.consultarEmpresa.clicked.connect(consultarempresas)
    TelaPrincipal.actionRelatorios_de_produtos.triggered.connect(acessarindicepesquisarprodutos)
    # TelaPrincipal.pesquisar.textChanged.connect(tentaracesar)
    telaDeLogin.lineEdit_2.setEchoMode(QtWidgets.QLineEdit.Password)# Inplementando campo senha
    # Conexão com o DB
    try:
        
        with open('Config\\config.json') as f:# Abrindo o arquivo que contem a string de conexao com o DB
            entrada = json.load(f)# lendo o json e armazenando em uma variavel
        
        conexao = mysql.connector.Connect(# string de conexao
            host=entrada["host"],
            user=entrada["user"],
            password=entrada["password"],
            database=entrada["database"],
            auth_plugin=entrada["auth_plugin"]
        )
        dataagoraservico = dt.datetime.now()
        logging.info("Iniciando conexão com o banco de dados")
    except Exception as erro:
        logging.exception(erro)
        QMessageBox.critical(TelaPrincipal, "Aviso", "{}".format(erro))

    consultas()
    tentaracesar()
    telaDeLogin.show()
    TelaPrincipal.stackedWidget.setCurrentIndex(1)
    app.exec()