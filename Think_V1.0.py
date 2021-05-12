# Esses são os frameworks importados para criaço do sistema
# temos o QTdesigner e também o mysql.conector para conexo com o MySQL
from pycep_correios import get_address_from_cep, WebService, exceptions
from PyQt5 import QtWidgets, uic
import mysql.connector
import xlwt
import pandas as pd
import win32com.client as win32
from pywintypes import com_error
import datetime as dt
import openpyxl




# Conexão com o banco de dados passando os parametros abaixo.
banco = mysql.connector.Connect(
    host="localhost",
    user="Weslei",
    password="0803",
    database="think",
    auth_plugin='mysql_native_password'
)


def chama_segunda_tela():

    primeira_tela.label_4.setText("")
    nome_usuario = primeira_tela.lineEdit.text()
    senha = primeira_tela.lineEdit_2.text()

    try:
        cursor = banco.cursor()
        cursor.execute("""select senha FROM
            usuarios WHERE nome ='{}'""".format(nome_usuario))
        senha_bd = cursor.fetchall()
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
        cursor.execute("""select idproduto, descricao, dt_entrada, preco,
         observacao, marca,
         referencia from produtos order by idproduto DESC limit 1200""")

        dados_lidos1 = cursor.fetchall()
        formulario.tableWidget.setRowCount(len(dados_lidos1))
        formulario.tableWidget.setColumnCount(7)
        for i in range(0, len(dados_lidos1)):
            for j in range(0, 7):
                formulario.tableWidget.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(dados_lidos1[i][j])))


        cursor = banco.cursor()
        cursor.execute("""select idvenda, nomecliente,
            tipo_negociacao_idtipo_negociacao,
            vendedores_idvendedor, data_venda, dat_venv_fatuura, nomeproduto,
            quantproduto, precoproduto, descproduto, vlr_total, id_tipopagamento,
            vezesdeparcelas, observacao, entrada_saida FROM vendas order by idvenda DESC limit 1000000;""")


        sql_vendas1 = cursor.fetchall()

        formulario.tableWidget_4.setRowCount(len(sql_vendas1))
        formulario.tableWidget_4.setColumnCount(15)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 15):
                formulario.tableWidget_4.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

        # Estou pegando a data do dia nessa variavél datavenda
        datadavenda = dt.datetime.now()
        # Estou setando o valor da variavél no objeto da data
        formulario.dateEdit_12.setDate(datadavenda)
        formulario.dateEdit_13.setDate(datadavenda)
        formulario.dateEdit_2.setDate(datadavenda)
        formulario.dateEdit_3.setDate(datadavenda)
        # formulario.dateEdit_4.setDate(datadavenda)
        # formulario.dateEdit_11.setDate(datadavenda)
        # formulario.dateEdit_8.setDate(datadavenda)
        # formulario.dateEdit_9.setDate(datadavenda)
        # formulario.dateEdit_10.setDate(datadavenda)
        # formulario.dateEdit_14.setDate(datadavenda)
        # formulario.dateEdit_15.setDate(datadavenda)
        # formulario.dateEdit_16.setDate(datadavenda)


    except IndexError as indexx:
        primeira_tela.label_4.setText("{}".format(indexx))
        return
    except mysql.connector.errors.InterfaceError as erro:
        aviso.show()
        aviso.textBrowser.setText("10061 Nenhuma conexão pôde ser feita porque \na máquina de destino as recusou ativamente{}".format(erro))
        return
    else:
        primeira_tela.label_4.setText("Dados de login incorretos!")
        return
    cursor.close()

def cadastrar_produtos():

    try:

        datadaentrada = dt.datetime.now()
        estoque = (formulario.lineEdit.text())
        descricao = str(formulario.lineEdit_2.text())
        preco = (formulario.lineEdit_3.text())
        ref = (formulario.lineEdit_4.text())
        observacao = str(formulario.lineEdit_9.text())
        marca = str(formulario.lineEdit_5.text())
        categoria = str(formulario.comboBox_8.currentText())
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



        cursor = banco.cursor()
        SQL_produtos = """insert into produtos
         (categorias_idcategoria, descricao, preco, observacao,marca,referencia, dt_entrada)
         VALUES  (%s,%s,%s,%s,%s,%s,%s)"""
        dados1 = (categoria), (descricao), (preco), (observacao), (marca), (ref), (datadaentrada)
        if not preco:
            aviso.show()
            aviso.textBrowser.setText("  Preencha os campos vazios EX: Preço.")
            return
        elif not descricao:
            aviso.show()
            aviso.textBrowser.setText("  Preencha os campos vazios EX:Descrição.")
            return
        cursor.execute(SQL_produtos, dados1)
        banco.commit()

        cursor.close()
        cursor = banco.cursor()
        if not estoque:
            formulario.lineEdit_7.setText("Preencha o campo Estoque")
            return
        cursor.callproc("PRC_EST", [estoque])
        banco.commit()
        cursor = banco.cursor()
        cursor.execute("""SELECT idproduto, descricao, dt_entrada, preco, observacao, marca,
         referencia from produtos order by idproduto DESC limit 1200""")
        sql_tprodu = cursor.fetchall()

        formulario.tableWidget.setRowCount(len(sql_tprodu))
        formulario.tableWidget.setColumnCount(7)

        for i in range(0, len(sql_tprodu)):
            for j in range(0, 7):
                formulario.tableWidget.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_tprodu[i][j])))

        cursor.close()

        formulario.lineEdit.setText("")
        formulario.lineEdit_2.setText("")
        formulario.lineEdit_3.setText("")
        formulario.lineEdit_4.setText("")
        formulario.lineEdit_5.setText("")
        formulario.lineEdit_9.setText("")

    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return

    except mysql.connector.errors.DatabaseError as erro:
        aviso.show()
        aviso.textBrowser.setText("{}".format(erro))
        return


def cadastrar_usuario():

    nome = tela_cadastro.lineEdit.text()
    email = tela_cadastro.lineEdit_2.text()
    senha = tela_cadastro.lineEdit_3.text()
    c_senha = tela_cadastro.lineEdit_4.text()

    if (senha == c_senha):
        try:

            cursor = banco.cursor()
            sql_user = """INSERT INTO usuarios (nome, email, senha)
            VALUES ('{}','{}','{}')""".format(nome, email, senha)
            cursor.execute(sql_user, dados_user)
            banco.commit()



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


def virificacep():

    try:
        cep = formulario.cepCliente.text()

        if not cep:
            return False
        else:
            aviso.textBrowser.setText("Sucesso na consulta do CEP.")


        endereco = get_address_from_cep(
            cep, webservice=WebService.APICEP)

    except exceptions.CEPNotFound as testa:
        aviso.show()
        aviso.textBrowser.setText("Cep Invalido\n{}".format(testa))
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
    try:
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

        if estadodocliente == "AC":
            estadodocliente = "1"
        elif estadodocliente == "AL":
            estadodocliente = "2"
        elif estadodocliente == "AM":
            estadodocliente = "3"
        elif estadodocliente == "AP":
            estadodocliente = "4"
        elif estadodocliente == "BA":
            estadodocliente = "5"
        elif estadodocliente == "CE":
            estadodocliente = "6"
        elif estadodocliente == "DF":
            estadodocliente = "7"
        elif estadodocliente == "ES":
            estadodocliente = "8"
        elif estadodocliente == "GO":
            estadodocliente = "9"
        elif estadodocliente == "MA":
            estadodocliente = "10"
        elif estadodocliente == "MG":
            estadodocliente = "11"
        elif estadodocliente == "MS":
            estadodocliente = "12"
        elif estadodocliente == "MT":
            estadodocliente = "13"
        elif estadodocliente == "PA":
            estadodocliente = "14"
        elif estadodocliente == "PB":
            estadodocliente = "15"
        elif estadodocliente == "PE":
            estadodocliente = "16"
        elif estadodocliente == "PI":
            estadodocliente = "17"
        elif estadodocliente == "PR":
            estadodocliente = "18"
        elif estadodocliente == "RJ":
            estadodocliente = "19"
        elif estadodocliente == "RN":
            estadodocliente = "20"
        elif estadodocliente == "RO":
            estadodocliente = "21"
        elif estadodocliente == "RR":
            estadodocliente = "22"
        elif estadodocliente == "RS":
            estadodocliente = "23"
        elif estadodocliente == "SC":
            estadodocliente = "24"
        elif estadodocliente == "SE":
            estadodocliente = "25"
        elif estadodocliente == "SP":
            estadodocliente = "26"
        elif estadodocliente == "TO":
            estadodocliente = "27"
        else:
            estadodocliente = "8"



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

        sql_cliente = """ insert into parceiros
        (bairros_cidades_idcidade, bairros_cidades_estados_idestado,
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
        banco.commit()
        cursor.close()

        aviso.show()
        aviso.textBrowser.setText("Cliente cadastrado com sucesso!!!")
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return


def vender_produto():

    try:
    	# dthoje = dt.datetime.now()
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
        entrada_saida = formulario.comboBox_3.currentText()
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

        if entrada_saida == '1 - Entrada':
        	entrada_saida = int(1)
        elif entrada_saida == '2 - Saida':
        	entrada_saida = int(2)


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

        cursor = banco.cursor()

        sql_vendas = """ INSERT INTO vendas (tipo_negociacao_idtipo_negociacao, vendedores_usuarios_idusuarios,
        vendedores_idvendedor, data_venda, vlr_total, nomecliente, nomeproduto,quantproduto, precoproduto,descproduto, id_tipopagamento, vezesdeparcelas, observacao, dat_venv_fatuura, entrada_saida)
        VALUES({}, {}, {}, '{}',{}, '{}', '{}', {}, {}, {}, '{}', {}, '{}', {}, {});""".format(
            tipoNegociacao, idusuario,
            iddovendedor, datadavenda,valortotaldavenda, nomeclientevenda, nomeprodutovenda,
            quantidadedoprodutovenda, precodoproduto, descontodoproduto, naturezavenda,
             vezesparcelas, obcervacaovenda, datadevencimento, entrada_saida)




        cursor.execute(sql_vendas)
        banco.commit()


        cursor = banco.cursor()
        cursor.execute("""select idvenda, nomecliente,
            tipo_negociacao_idtipo_negociacao,
            vendedores_idvendedor, data_venda, dat_venv_fatuura, nomeproduto,
            quantproduto, precoproduto, descproduto, vlr_total, id_tipopagamento,
            vezesdeparcelas, observacao, entrada_saida FROM vendas order by idvenda DESC limit 1000000;""")


        sql_vendas1 = cursor.fetchall()

        formulario.tableWidget_4.setRowCount(len(sql_vendas1))
        formulario.tableWidget_4.setColumnCount(15)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 15):
                formulario.tableWidget_4.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

        formulario.lineEdit_19.setText("")
        # formulario.lineEdit_10.setText("")
        formulario.lineEdit_17.setText("")
        formulario.lineEdit_14.setText("")
        formulario.lineEdit_15.setText("")
        formulario.lineEdit_16.setText("")

    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return




# função para deletar um produto
def deletarVendas():

    try:

        cursor = banco.cursor()
        deletarvenda = formulario.lineEdit_11.text()
        if not deletarvenda:
            return False

        deletarvendasql = """DELETE FROM vendas WHERE idvenda = {}""".format(deletarvenda)

        cursor.execute(deletarvendasql)
        banco.commit()

        aviso.show()
        aviso.textBrowser.setText('Produto deletado com exito.')
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return
        cursor.close()

# função que abre a tela de cadastro
def abre_tela_cadastro():
    # o metodo show() é usado para chamar a tela
    tela_cadastro.show()

def gerarelatorio():

    try:

        df = pd.read_sql('select * from Vendas', banco)
        salvar = QtWidgets.QFileDialog.getSaveFileName()[0]
        df.to_excel(salvar + '.xlsx', index=False)
        tela_progresso.show()
        tela_progresso.progressBar.setValue(100)
    except ValueError as er:
        return False
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))


def gerarelatorio_produtos():
    try:

        df = pd.read_sql('select * from produtos', banco)
        salvar = QtWidgets.QFileDialog.getSaveFileName()[0]
        df.to_excel(salvar + '.xlsx', index=False)
        tela_progresso.show()
        tela_progresso.progressBar.setValue(100)
    except ValueError as er:
        return False
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))

def gerarelatorioEstoque():
    try:

        df = pd.read_sql('select * from estoque', banco)
        salvar = QtWidgets.QFileDialog.getSaveFileName()[0]
        df.to_excel(salvar + '.xlsx', index=False)
        tela_progresso.show()
        tela_progresso.progressBar.setValue(100)
    except ValueError as er:
        return False
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))

def geraRelatorioVendasEntSaida():
    try:

        datavenda1 = formulario.dateEdit_12.text()
        datavenda1 = dt.datetime.strptime(datavenda1, '%d/%m/%Y')
        datavenda1 = datavenda1.strftime('%Y-%m-%d')

        datavenda2 = formulario.dateEdit_13.text()
        datavenda2 = dt.datetime.strptime(datavenda2, '%d/%m/%Y')
        datavenda2 = datavenda2.strftime('%Y-%m-%d')
        ent_sai = formulario.comboBox.currentText()

        if ent_sai == '1 - Entrada':
            ent_sai = int(1)
        elif ent_sai == '2 - Saida':
            ent_sai =  int(2)

        df = pd.read_sql("""select idvenda, tipo_negociacao_idtipo_negociacao, vendedores_usuarios_idusuarios,
            vendedores_idvendedor, data_venda, vlr_total, nomecliente, nomeproduto,
            quantproduto, precoproduto, descproduto, id_tipopagamento, vezesdeparcelas,
            observacao, dat_venv_fatuura, entrada_saida
            FROM vendas where entrada_saida = {} and data_venda >= '{}' and data_venda <= '{}'order
            by idvenda DESC limit 1000000;""".format(ent_sai,datavenda1, datavenda2), banco)

        salvar = QtWidgets.QFileDialog.getSaveFileName()[0]
        df.to_excel(salvar + '.xlsx', index=False)
        tela_progresso.show()
        tela_progresso.progressBar.setValue(100)
    except ValueError as er:
        return False
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))


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
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
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

        <p>Olá Bom Dia!</p>

        <p>A quantidade de produtos em estoque é {}.</p>

        <p>Vendemos {} produtos.</p>

        <p>O valor atual do produto é {}.</p>

        <p>{}.</p>

        <p>Att, .</p>

        <p>Sistema Think Gestão em microempresas.</p>

        """.format(faturamento, vendas, ticket, cdialogo)
        if not email.to:
            aviso.show()
            aviso.textBrowser.setText('Microsoft Outlook Precisamos saber para quem enviar isto. Verifique se voc� inseriu pelo menos um nome.')
            return
        email.Send()
        aviso.show()
        aviso.textBrowser.setText('Email enviado para\n{}'.format(email2))

        formulario.lineEdit_29.setText('')
        formulario.lineEdit_30.setText('')
        formulario.lineEdit_31.setText('')
        formulario.lineEdit_32.setText('')
        formulario.lineEdit_33.setText('')
        formulario.lineEdit_34.setText('')
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return

def vendas_realizadas():
    # Aqui é feito a converssão das datas
    datavenda1 = formulario.dateEdit_3.text()
    datavenda1 = dt.datetime.strptime(datavenda1, '%d/%m/%Y')
    datavenda1 = datavenda1.strftime('%Y-%m-%d')

    datavenda2 = formulario.dateEdit_2.text()
    datavenda2 = dt.datetime.strptime(datavenda2, '%d/%m/%Y')
    datavenda2 = datavenda2.strftime('%Y-%m-%d')

    try:

        cursor = banco.cursor()
        cursor.execute("""select idvenda, nomecliente,
                tipo_negociacao_idtipo_negociacao,
                vendedores_idvendedor, data_venda, dat_venv_fatuura, nomeproduto,
                quantproduto, precoproduto, descproduto, vlr_total, id_tipopagamento,
                vezesdeparcelas, observacao, entrada_saida FROM vendas
                where vezesdeparcelas > 1 and data_venda >= '{}' and data_venda <= '{}' order by idvenda DESC limit 1000000;""".format(datavenda1, datavenda2))


        sql_vendas1 = cursor.fetchall()

        formulario.tableWidget_5.setRowCount(len(sql_vendas1))
        formulario.tableWidget_5.setColumnCount(15)

        for i in range(0, len(sql_vendas1)):
            for j in range(0, 15):
                formulario.tableWidget_5.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sql_vendas1[i][j])))

        cursor.close()

    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return
def teladevendas():
    try:
        datavenda1 = formulario.dateEdit_12.text()
        datavenda1 = dt.datetime.strptime(datavenda1, '%d/%m/%Y')
        datavenda1 = datavenda1.strftime('%Y-%m-%d')

        datavenda2 = formulario.dateEdit_13.text()
        datavenda2 = dt.datetime.strptime(datavenda2, '%d/%m/%Y')
        datavenda2 = datavenda2.strftime('%Y-%m-%d')
        ent_sai = formulario.comboBox.currentText()

        if ent_sai == '1 - Entrada':
            ent_sai = int(1)
        elif ent_sai == '2 - Saida':
            ent_sai =  int(2)

        cursor = banco.cursor()
        cursor.execute("""select idvenda, nomecliente,
                tipo_negociacao_idtipo_negociacao,
                vendedores_idvendedor, data_venda, dat_venv_fatuura, nomeproduto,
                quantproduto, precoproduto, descproduto, vlr_total, id_tipopagamento,
                vezesdeparcelas, observacao, entrada_saida
            FROM vendas where entrada_saida = {} and data_venda >= '{}' and data_venda <= '{}'order by idvenda DESC limit 1000000;""".format(ent_sai,datavenda1, datavenda2))
        sqlVendas = cursor.fetchall()
        formulario.tableWidget_7.setRowCount(len(sqlVendas))
        formulario.tableWidget_7.setColumnCount(15)

        for i in range(0, len(sqlVendas)):
            for j in range(0, 15):
                formulario.tableWidget_7.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVendas[i][j])))
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return




def vendasAvista():
    # Aqui é feito a converssão das datas
    categoriaDavenda = formulario.comboBox_5.currentText()
    datavenda1 = formulario.dateEdit_3.text()
    datavenda1 = dt.datetime.strptime(datavenda1, '%d/%m/%Y')
    datavenda1 = datavenda1.strftime('%Y-%m-%d')

    datavenda2 = formulario.dateEdit_2.text()
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
        return False


    try:
        cursor =  banco.cursor()
        cursor.execute(""" select idvenda, nomecliente,
                tipo_negociacao_idtipo_negociacao,
                vendedores_idvendedor, data_venda, dat_venv_fatuura, nomeproduto,
                quantproduto, precoproduto, descproduto, vlr_total, id_tipopagamento,
                vezesdeparcelas, observacao, entrada_saida FROM vendas where vezesdeparcelas <= 1
                and data_venda >= '{}' and data_venda <= '{}' order by idvenda DESC limit 1000000;""".format(datavenda1, datavenda2))
        sqlVendasAvista = cursor.fetchall()
        formulario.tableWidget_5.setRowCount(len(sqlVendasAvista))
        formulario.tableWidget_5.setColumnCount(15)

        for i in range(0, len(sqlVendasAvista)):
            for j in range(0, 15):
                formulario.tableWidget_5.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVendasAvista[i][j])))
        cursor.close()
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return


def verificarProduto():
    try:
        codigodeverificacao = formulario.lineEdit_45.text()
        referenciaparaverificacao = formulario.lineEdit_44.text()
        cursor = banco.cursor()
        cursor.execute("""select idproduto, descricao, dt_entrada, preco,
         observacao, marca, referencia from produtos where referencia = '{}' order by idproduto DESC limit 1200""".format(referenciaparaverificacao))


        sqlVerificacaoProduto = cursor.fetchall()
        formulario.tableWidget.setRowCount(len(sqlVerificacaoProduto))
        formulario.tableWidget.setColumnCount(7)

        for i in range(0, len(sqlVerificacaoProduto)):
            for j in range(0, 7):
                formulario.tableWidget.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVerificacaoProduto[i][j])))

        cursor = banco.cursor()
        if not codigodeverificacao:
            return False
        cursor.execute("""select * from estoque where produtos_idproduto={};""".format(codigodeverificacao))

        sqlVerificacaoProduto = cursor.fetchall()
        formulario.tableWidget_2.setRowCount(len(sqlVerificacaoProduto))
        formulario.tableWidget_2.setColumnCount(5)

        for i in range(0, len(sqlVerificacaoProduto)):
            for j in range(0, 5):
                formulario.tableWidget_2.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVerificacaoProduto[i][j])))
        cursor.close()
        formulario.lineEdit_45.setText("")
        formulario.lineEdit_44.setText("")
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return


def deletarProduto():
    try:
        codigoVenda = formulario.lineEdit_21.text()
        codigoDoproduto = formulario.lineEdit_23.text()
        if not codigoVenda:
            return False
        cursor = banco.cursor()
        cursor.execute("""delete from estoque where produtos_idproduto ={}""".format(codigoVenda))
        banco.commit()
        cursor.execute("""delete from produtos where idproduto={}""".format(codigoDoproduto))
        banco.commit()
        cursor.execute("""select * from estoque;""")

        sqlVerificacaoProduto = cursor.fetchall()
        formulario.tableWidget_2.setRowCount(len(sqlVerificacaoProduto))
        formulario.tableWidget_2.setColumnCount(5)

        for i in range(0, len(sqlVerificacaoProduto)):
            for j in range(0, 5):
                formulario.tableWidget_2.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVerificacaoProduto[i][j])))

        referenciaProduto = formulario.lineEdit_23.text()
        if not referenciaProduto:
            return False



        formulario.lineEdit_21.setText("")
        formulario.lineEdit_23.setText("")
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return


def atualizarEstoque():
    try:
        codigoDoproduto =  formulario.lineEdit_46.text()
        quantidadeAtualdoEstoque = formulario.lineEdit_47.text()
        cursor =  banco.cursor()
        if not codigoDoproduto:
            return False
        elif not quantidadeAtualdoEstoque:
            return False
        cursor.execute("""update estoque set estoque = {} where produtos_idproduto={};""".format(quantidadeAtualdoEstoque, codigoDoproduto))
        banco.commit()

        codigodeverificacao = formulario.lineEdit_46.text()
        if not codigodeverificacao:
            return False
        cursor.execute("""select * from estoque where produtos_idproduto={};""".format(codigodeverificacao))

        sqlVerificacaoProduto = cursor.fetchall()
        formulario.tableWidget_2.setRowCount(len(sqlVerificacaoProduto))
        formulario.tableWidget_2.setColumnCount(5)

        for i in range(0, len(sqlVerificacaoProduto)):
            for j in range(0, 5):
                formulario.tableWidget_2.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVerificacaoProduto[i][j])))

        cursor.close()
        formulario.lineEdit_46.setText("")
        formulario.lineEdit_47.setText("")

    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return
def pesquisarProduto():
    try:
        pesquisar = formulario.lineEdit_12.text()
        cursor = banco.cursor()
        cursor.execute("""select idproduto, descricao, dt_entrada, preco,
         observacao, marca, referencia from produtos where descricao like '{}'""".format(pesquisar))

        sqlVerificacaoProduto = cursor.fetchall()
        formulario.tableWidget.setRowCount(len(sqlVerificacaoProduto))
        formulario.tableWidget.setColumnCount(7)

        for i in range(0, len(sqlVerificacaoProduto)):
            for j in range(0, 7):
                formulario.tableWidget.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVerificacaoProduto[i][j])))

        cursor.close()
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return


def consultaprodutoParavenda():
    try:
        pesquisar = formulario.lineEdit_43.text()
        cursor = banco.cursor()
        cursor.execute("""select idproduto, descricao, dt_entrada, preco,
         observacao, marca, referencia from produtos where descricao like '{}'""".format(pesquisar))

        sqlVerificacaoProduto = cursor.fetchall()
        formulario.tableWidget_3.setRowCount(len(sqlVerificacaoProduto))
        formulario.tableWidget_3.setColumnCount(7)

        for i in range(0, len(sqlVerificacaoProduto)):
            for j in range(0, 7):
                formulario.tableWidget_3.setItem(
                    i, j, QtWidgets.QTableWidgetItem(str(sqlVerificacaoProduto[i][j])))

        cursor.close()
    except Exception as e:
        aviso.show()
        aviso.textBrowser.setText('{}'.format(e))
        return


# Conectores e operadores
# App recebe os objetos QT
app = QtWidgets.QApplication([])
#app.setStyle (QtWidgets.QStyleFactory.create ('Fusion'))
app.setStyle ( 'windowsvista' )
# As variavés abaixo recebem e carregam as telas com o metodo loadUi
primeira_tela = uic.loadUi("Telas\\teladelogin.ui")
tela_cadastro = uic.loadUi("Telas\\tela_cadastro.ui")
formulario = uic.loadUi("Telas\\segunda_tela.ui")
aviso = uic.loadUi("Telas\\avisosnovos.ui")
tela_progresso = uic.loadUi("Telas\\barradeprogreço.ui")
# Abaixo conectamos um objeto ao outro para que possão interagir conforme
# nosso projeto, ex: A variavél primeira_tela recebe o botão e o botão
# recebe o metodo clicKed com o metodo connect para conectar a função
# que é passada dentro como parametro
formulario.salvarCliente.clicked.connect(cadcliente)
# formulario.salvarCliente.setToolTip("Salvar novo cliente")
primeira_tela.pushButton.clicked.connect(chama_segunda_tela)
formulario.pushButton_9.clicked.connect(deletarVendas)
formulario.enviaemail.clicked.connect(enviaremail)
formulario.actionRELATORIO_DE_VENDAS.triggered.connect(gerarelatorio)

# Aqui usamos o metodo Password para indicar que esse campo vai receber
# um valor do tipo senha e recebera o '*' como padão de segurança.
primeira_tela.lineEdit_2.setEchoMode(QtWidgets.QLineEdit.Password)
# Os botões que Abrem as telas
formulario.pushButton_8.clicked.connect(virificacep)
# formulario.pushButton_8.setToolTip("Consultar Cep")
primeira_tela.pushButton_2.clicked.connect(abre_tela_cadastro)
tela_cadastro.pushButton.clicked.connect(cadastrar_usuario)
formulario.pushButton.clicked.connect(cadastrar_produtos)
formulario.pushButton_4.clicked.connect(vender_produto)
formulario.pushButton_10.clicked.connect(vendas_realizadas)
formulario.geraexcel.clicked.connect(gerarelatorio)
formulario.geraexcel_3.clicked.connect(gerarelatorioEstoque)
formulario.geraexcel_2.clicked.connect(gerarelatorio_produtos)
formulario.geraexcel_4.clicked.connect(geraRelatorioVendasEntSaida)
formulario.enviaemail_2.clicked.connect(enviaremailprodutos)
# formulario.pushButton_16.clicked.connect(deletarProduto)
formulario.pushButton_3.clicked.connect(teladevendas)
formulario.pushButton_15.clicked.connect(vendasAvista)
formulario.pushButton_14.clicked.connect(vendasAvista)
formulario.pushButton_2.clicked.connect(verificarProduto)
formulario.pushButton_6.clicked.connect(deletarProduto)
formulario.pushButton_11.clicked.connect(atualizarEstoque)
formulario.pushButton_7.clicked.connect(pesquisarProduto)
formulario.pushButton_20.clicked.connect(consultaprodutoParavenda)
# Aqui eu pego os dados em um select com pandas para adicionar nos comboBox.
df = pd.read_sql('select nome from usuarios', banco)
formulario.comboBox_7.addItems(df['nome'])
formulario.comboBox_2.addItems(df['nome'])
##########################################################################
df = pd.read_sql('select descricao from tipo_negociacao', banco)
formulario.tipoNegociacao.addItems(df['descricao'])
##########################################################################
df = pd.read_sql('select descricao from categorias', banco)
formulario.comboBox_8.addItems(df['descricao'])
##########################################################################
df = pd.read_sql('select descricao from tipo_pagamento;', banco)
formulario.naturezadavenda.addItems(df['descricao'])
##########################################################################
df = pd.read_sql('select tipo_cliente from cat_cliente;', banco)
formulario.catCliente.addItems(df['tipo_cliente'])
##########################################################################
formulario.vencimentoparcelado.addItems([ '1', '2', '3', '4', '5',
'6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18',
'19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30'  ])
formulario.vezesdeparcelas.addItems(['1X',
    '2X', '3X','4X', '5X', '6X', '7X', '8X', '9X', '10X'
    ])
##########################################################################
formulario.comboBox_3.addItems(['1 - Entrada', '2 - Saida'])
formulario.comboBox.addItems(['1 - Entrada', '2 - Saida'])

df = pd.read_sql('select descricao from tipo_pagamento;', banco)
formulario.comboBox_4.addItems(df['descricao'])

df = pd.read_sql('select descricao from tipo_pagamento;', banco)
formulario.comboBox_5.addItems(df['descricao'])

# Esse premeiro show() é aonde o sistema começa.
primeira_tela.show()
# formulario.show()

app.exec()
