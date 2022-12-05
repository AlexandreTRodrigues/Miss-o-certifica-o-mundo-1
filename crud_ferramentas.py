import sqlite3

#criar classe para interagir o cadastro de ferramentas com banco de dados

class crud_ferramentas:
    def __init__(self):
        pass

    def abrirConexaoFerramentas(self):
        self.conexaoFerramentas = sqlite3.connect('ferramentas_clientes.db')
        print('Conexão com BD obtida com sucesso')

#-----------------------------------------------------------------------------------------------------
#cadastrar ferramenta

    def cadastrarFerramenta(self, descriçao, fabricante, voltagem, part_number, tamanho, unidade_medida, tipo, material, tempo_maximo):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        insert_sql = '''INSERT INTO cadastro_ferramentas('descrição', 'fabricante', 'voltagem', 'part_number',
        'tamanho', 'unidade_medida', 'tipo', 'material', 'tempo_maximo')
        VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)'''
        grava_insert = (descriçao, fabricante, voltagem, part_number, tamanho,
                        unidade_medida, tipo, material, tempo_maximo)
        cursor.execute(insert_sql, grava_insert)

        self.conexaoFerramentas.commit()
        cursor.close()
        self.conexaoFerramentas.close()

#-----------------------------------------------------------------------------------------------------
#selecionar ferramentas do BD

    def selecionarFerramenta(self):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        seleciona_sql = '''SELECT * FROM cadastro_ferramentas'''
        cursor.execute(seleciona_sql)
        ferramentas = cursor.fetchall()

        cursor.close()
        self.conexaoFerramentas.close()
        return ferramentas


# -----------------------------------------------------------------------------------------------------
# deletar ferramenta

    def deletarFerramenta(self, id):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        comando_sql = '''DELETE from cadastro_ferramentas WHERE id = ?'''
        cursor.execute(comando_sql, (id, ))
        self.conexaoFerramentas.commit()
        cursor.close()
        self.conexaoFerramentas.close()

# -----------------------------------------------------------------------------------------------------
# atualizar ferramenta

    def atualizaFerramenta(self, id, descriçao, fabricante, voltagem, part_number, tamanho, unidade_medida, tipo, material, tempo_maximo):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        seleciona_tudo_sql = '''SELECT * FROM cadastro_ferramentas WHERE id = ?'''
        cursor.execute(seleciona_tudo_sql, (id, ))
        atualiza_sql = '''UPDATE cadastro_ferramentas SET descrição = ?, fabricante = ?,
        voltagem = ?, part_number = ?, tamanho = ?, unidade_medida = ?, tipo = ?, material = ?,
        tempo_maximo = ? WHERE id = ?'''
        cursor.execute(atualiza_sql, (id, descriçao, fabricante, voltagem, part_number, tamanho, unidade_medida, tipo, material, tempo_maximo))
        self.conexaoFerramentas.commit()
        count = cursor.rowcount
        print(count, 'Registro atualizado com sucesso!')
        cursor.close()
        self.conexaoFerramentas.close()

# -----------------------------------------------------------------------------------------------------
# cadastrar técnico

    def cadastrarTecnico(self, cpf, nome, telefone, turno, equipe):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        dados_tecnico = (cpf, nome, telefone, turno, equipe)
        comando = '''INSERT INTO cadastro_tecnico(cpf, nome, telefone, turno, equipe) VALUES(?, ?, ?, ?, ?)'''
        cursor.execute(comando, dados_tecnico)
        self.conexaoFerramentas.commit()
        cursor.close()
        self.conexaoFerramentas.close()

# -----------------------------------------------------------------------------------------------------
# selecionar técnico

    def selecionaTecnico(self):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        select_dados = '''SELECT * FROM cadastro_tecnico'''
        cursor.execute(select_dados)
        tecnico = cursor.fetchall()
        cursor.close()
        self.conexaoFerramentas.close()
        return  tecnico
# -----------------------------------------------------------------------------------------------------
# deletar técnico

    def deletaTecnico(self, cpf):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        comando = '''DELETE FROM cadastro_tecnico WHERE cpf = ?'''
        cursor.execute(comando, (cpf, ))
        self.conexaoFerramentas.commit()
        cursor.close()
        self.conexaoFerramentas.close()

# -----------------------------------------------------------------------------------------------------
# atualizar técnico

    def atualizaTecnico(self, cpf, nome, telefone, turno, equipe):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        select_sql = '''SELECT * FROM cadastro_tecnico WHERE cpf = ?'''
        cursor.execute(select_sql, (cpf, ))
        atualiza_sql = '''UPDATE cadastro_tecnico SET nome = ?, telefone = ?, turno = ?, equipe = ? WHERE cpf = ?'''
        cursor.execute(atualiza_sql, (cpf, nome, telefone, turno, equipe))
        self.conexaoFerramentas.commit()
        cursor.close()
        self.conexaoFerramentas.close()

# -----------------------------------------------------------------------------------------------------
# cadastrar solicitação

    def cadastraSolicitacao(self, id_ferramenta, descri_solic, data_ret, hora_ret, data_dev, hora_dev, nome_tec):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        cadastra_sql = '''INSERT INTO cadastro_solicitação(id_ferramenta, descri_solic, data_ret, hora_ret,
        data_dev, hora_dev, nome_tec) VALUES(?,?,?,?,?,?,?)'''
        grava_cadastro = (id_ferramenta, descri_solic, data_ret, hora_ret, data_dev, hora_dev, nome_tec)
        cursor.execute(cadastra_sql, grava_cadastro)
        self.conexaoFerramentas.commit()
        cursor.close()
        self.conexaoFerramentas.close()

# -----------------------------------------------------------------------------------------------------
# deletar solicitação

    def deletarSolicitacao(self, id_ferramenta):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        deleta_sql = '''DELETE FROM cadastro_solicitação WHERE id_ferramenta = ?'''
        cursor.execute(deleta_sql, (id_ferramenta, ))
        self.conexaoFerramentas.commit()
        cursor.close()
        self.conexaoFerramentas.close()

# -----------------------------------------------------------------------------------------------------
# selecionar solicitação

    def selecionaSolicitacao(self):
        self.abrirConexaoFerramentas()
        cursor = self.conexaoFerramentas.cursor()
        seleciona_sql = '''SELECT * FROM cadastro_solicitação'''
        cursor.execute(seleciona_sql)
        solicitacao = cursor.fetchall()
        cursor.close()
        self.conexaoFerramentas.close()
        return solicitacao
