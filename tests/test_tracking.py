"""Regressões de distribuição e atualização, usando apenas dados fictícios."""
from pathlib import Path
import hashlib
import sys
import tempfile
import unittest
from collections import Counter
from datetime import datetime
from zipfile import ZipFile

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / 'Programa'))
from distribuir_cases import (carregar_base, carregar_exclusoes, separar_exclusoes,
                               distribuir, salvar_resultado, load_workbook, ErroBase)
from atualizar_painel import atualizar


class TrackingTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.pasta = Path(self.tmp.name)
        self.base = RAIZ / 'exemplos/Base_Ficticia_Cases.xlsx'
        self.nomes = ['Ana', 'Bruno', 'Carla', 'Diego', 'Elisa']

    def tearDown(self):
        self.tmp.cleanup()

    def gerar(self, espera=True):
        cases, aba = carregar_base(self.base)
        numeros = carregar_exclusoes(RAIZ / 'exemplos/exclusoes_demo.txt') if espera else []
        cases, excluidos, ausentes = separar_exclusoes(cases, numeros)
        distribuir(cases, self.nomes)
        return salvar_resultado(cases, self.nomes, self.base, aba, self.pasta, excluidos, ausentes)

    def test_distribuicao_completa_equilibrada(self):
        cases, _ = carregar_base(self.base)
        distribuir(cases, self.nomes)
        self.assertEqual(Counter(c.responsavel for c in cases),
                         dict(zip(self.nomes, [9, 9, 8, 8, 8])))
        self.assertEqual([c.abertura for c in cases], sorted(c.abertura for c in cases))

    def test_espera_e_numeros_ausentes(self):
        cases, _ = carregar_base(self.base)
        cases, espera, ausentes = separar_exclusoes(cases, carregar_exclusoes(RAIZ / 'exemplos/exclusoes_demo.txt'))
        distribuir(cases, self.nomes)
        self.assertEqual(len(cases), 40)
        self.assertEqual({c.numero for c in espera}, {'00090041', '00090042'})
        self.assertEqual(ausentes, ['00099999'])
        self.assertEqual(Counter(c.responsavel for c in cases), dict.fromkeys(self.nomes, 8))

    def test_planilha_com_sla_e_botao(self):
        arquivo = self.gerar()
        self.assertTrue(arquivo.name.startswith('Tracking_'))
        w = load_workbook(arquivo)
        try:
            self.assertEqual(w.sheetnames, ['Gestor', 'Tracking', 'Resumo', 'SLA'])
            self.assertEqual(w['Tracking'].max_row, 41)
            self.assertEqual({c.value for c in w['Tracking']['D'][1:]}, {'Pendente'})
            self.assertEqual(w['SLA'].tables['DetalheSLA'].ref, 'A7:H47')
            self.assertTrue(all(w['SLA'].cell(r, 11).value is None for r in range(10, 110)))
            self.assertEqual(w['SLA']['K3'].value, 4)
            self.assertTrue(w.calculation.fullCalcOnLoad)
            self.assertIn('WORKDAY', w['SLA']['E8'].value)
            self.assertEqual(w['Gestor'].row_breaks.brk[0].id, 28)
        finally:
            w.close()
        with ZipFile(arquivo) as z:
            self.assertIn('xl/vbaProject.bin', z.namelist())
            self.assertIn('GerarResumoPDF', z.read('xl/drawings/drawing1.xml').decode())

    def preparar_atualizacao(self, extensao):
        w = load_workbook(self.gerar())
        w['Tracking']['D2'] = 'Escalado'
        w['Tracking']['E2'] = 'Apoio'
        w['Tracking']['F2'] = 'Escalado para: Equipe Financeira - aguardando retorno.'
        w['Tracking']['D3'] = 'Concluído'
        w['SLA']['K10'] = datetime(2026, 9, 28)
        w['SLA']['K109'] = datetime(2026, 12, 25)
        origem = self.pasta / ('entrada' + extensao)
        if extensao == '.xlsx':
            w.save(origem)
        else:
            from incluir_botao_pdf import incluir_botao
            intermediario = self.pasta / 'entrada.part'
            w.save(intermediario)
            incluir_botao(intermediario, origem)
        antes = [[c.value for c in linha] for linha in w['Tracking']]
        resumo = {c.coordinate: c.value for linha in w['Resumo'] for c in linha if c.value is not None}
        w.close()
        digest = hashlib.sha256(origem.read_bytes()).hexdigest()
        saida = atualizar(origem, self.pasta / ('nova' + extensao))
        self.assertEqual(hashlib.sha256(origem.read_bytes()).hexdigest(), digest)
        novo = load_workbook(saida)
        try:
            self.assertEqual([[c.value for c in linha] for linha in novo['Tracking']], antes)
            self.assertEqual({c.coordinate: c.value for linha in novo['Resumo'] for c in linha if c.value is not None}, resumo)
            self.assertEqual(novo['SLA']['K10'].value, datetime(2026, 9, 28))
            self.assertEqual(novo['SLA']['K109'].value, datetime(2026, 12, 25))
        finally:
            novo.close()
        return saida

    def test_atualizacao_xlsx_preserva_atendimento(self):
        self.preparar_atualizacao('.xlsx')

    def test_atualizacao_xlsm_preserva_atendimento(self):
        self.preparar_atualizacao('.xlsm')

    def test_atualizacao_repetida_preserva_feriados(self):
        origem = self.preparar_atualizacao('.xlsm')
        saida = atualizar(origem, self.pasta / 'segunda')
        w = load_workbook(saida)
        self.assertEqual(w['SLA']['K109'].value, datetime(2026, 12, 25))
        w.close()

    def test_atualizacao_nao_sobrescreve_saida(self):
        origem = self.gerar()
        pasta = self.pasta / 'destino'
        atualizar(origem, pasta)
        with self.assertRaises(ErroBase):
            atualizar(origem, pasta)

    def test_base_rejeita_case_duplicado(self):
        w = load_workbook(self.base)
        s = w.active
        # A base de exemplo tem os cabeçalhos na primeira linha.
        s.cell(3, 1, s.cell(2, 1).value)
        arquivo = self.pasta / 'duplicada.xlsx'
        w.save(arquivo)
        w.close()
        with self.assertRaisesRegex(ErroBase, 'duplicado'):
            carregar_base(arquivo)


if __name__ == '__main__':
    unittest.main()
