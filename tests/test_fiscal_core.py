from pathlib import Path
import unittest

from fiscal_core import (
    ParametrosSimulacao,
    carregar_base,
    simular_creditos,
    simular_servicos,
)


ROOT = Path(__file__).resolve().parents[1]


class FiscalCoreTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.servicos, cls.contas, cls.fornecedores = carregar_base(ROOT / "ARQUIVOS")

    def test_le_todos_os_formatos_de_servicos(self):
        self.assertEqual(len(self.servicos), 718)
        self.assertEqual(
            set(self.servicos["formato"]),
            {"XML municipal", "XML NFS-e nacional", "TXT fixo municipal"},
        )

    def test_le_quatro_relatorios_de_contas(self):
        self.assertEqual(self.contas["arquivo"].nunique(), 4)
        self.assertGreater(len(self.contas), 100)
        self.assertGreater(self.contas["valor_base_credito"].sum(), 0)

    def test_reducao_de_saude(self):
        resultado = simular_servicos(self.servicos.head(1), ParametrosSimulacao())
        self.assertAlmostEqual(resultado.iloc[0]["aliquota_ibs_efetiva_pct"], 0.04)
        self.assertAlmostEqual(resultado.iloc[0]["aliquota_cbs_efetiva_pct"], 3.40)

    def test_retencao_so_quando_destacada(self):
        resultado = simular_servicos(self.servicos, ParametrosSimulacao())
        sem_ir = resultado[~resultado["ir_destacado"]]
        sem_csll = resultado[~resultado["csll_destacado"]]
        self.assertTrue((sem_ir["ir_2027"] == 0).all())
        self.assertTrue((sem_csll["csll_2027"] == 0).all())

    def test_credito_simples_e_reduzido(self):
        resultado = simular_creditos(self.contas, ParametrosSimulacao())
        simples = resultado[resultado["simples_nacional"].eq(True)]
        self.assertFalse(simples.empty)
        self.assertTrue((simples["aliquota_credito_cbs_pct"] <= 0.9258 + 1e-9).all())


if __name__ == "__main__":
    unittest.main()
