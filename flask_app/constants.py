# Sexo
SEXO_MASCULINO = 'M'
SEXO_FEMININO = 'F'
SEXO_OUTRO = 'O'

SEXO_CHOICES = [
    (SEXO_MASCULINO, 'Masculino'),
    (SEXO_FEMININO, 'Feminino'),
    (SEXO_OUTRO, 'Outro'),
]

# Status
STATUS_CIVIL = 'CIVIL'
STATUS_MILITAR_ATIVA = 'MILITAR DA ATIVA'
STATUS_MILITAR_RESERVA = 'MILITAR DA RESERVA'
STATUS_PENSIONISTA = 'PENSIONISTA'
STATUS_DEP_DESACOMPANHADO = 'DEP DESACOMPANHADO'
STATUS_DEP_ACOMPANHADO = 'DEP ACOMPANHADO'

STATUS_CHOICES = [
    (STATUS_CIVIL, 'CIVIL'),
    (STATUS_MILITAR_ATIVA, 'MILITAR DA ATIVA'),
    (STATUS_MILITAR_RESERVA, 'MILITAR DA RESERVA'),
    (STATUS_PENSIONISTA, 'PENSIONISTA'),
    (STATUS_DEP_DESACOMPANHADO, 'DEP DESACOMPANHADO'),
]

# Graduação
GRADUACAO_GEN = 'GEN'
GRADUACAO_CEL = 'CEL'
GRADUACAO_TC = 'TC'
GRADUACAO_MAJ = 'MAJ'
GRADUACAO_CAP = 'CAP'
GRADUACAO_1_TEN = '1º TEN'
GRADUACAO_2_TEN = '2º TEN'
GRADUACAO_ASP = 'ASP'
GRADUACAO_SO = 'SO'
GRADUACAO_ST = 'ST'
GRADUACAO_1_SGT = '1º SGT'
GRADUACAO_2_SGT = '2º SGT'
GRADUACAO_3_SGT = '3º SGT'
GRADUACAO_CIVIL = 'CIVIL'

GRADUACAO_CHOICES = [
    (GRADUACAO_GEN, 'GEN'),
    (GRADUACAO_CEL, 'CEL'),
    (GRADUACAO_TC, 'TC'),
    (GRADUACAO_MAJ, 'MAJ'),
    (GRADUACAO_CAP, 'CAP'),
    (GRADUACAO_1_TEN, '1º TEN'),
    (GRADUACAO_2_TEN, '2º TEN'),
    (GRADUACAO_ASP, 'ASP'),
    (GRADUACAO_SO, 'SO'),
    (GRADUACAO_ST, 'ST'),
    (GRADUACAO_1_SGT, '1º SGT'),
    (GRADUACAO_2_SGT, '2º SGT'),
    (GRADUACAO_3_SGT, '3º SGT'),
    (GRADUACAO_CIVIL, 'CIVIL'),
]

# Tipo
TIPO_CASAL = 'Casal'
TIPO_SOLTEIRO = 'Solteiro'
TIPO_OUTRO = 'Outro'

TIPO_CHOICES = [
    (TIPO_CASAL, 'Casal'),
    (TIPO_SOLTEIRO, 'Solteiro'),
    (TIPO_OUTRO, 'Outro'),
]

# Vínculo
VINCULO_CONJUGE = 'Cônjuge'
VINCULO_FILHO_ATE_6 = 'Filho até 6 anos'
VINCULO_FILHO_7_A_10 = 'Filho de 7 a 10 anos'
VINCULO_FILHO_11_A_23 = 'Filho de 11 a 23 anos'
VINCULO_FILHO_ACIMA_23 = 'Filho acima de 23 anos'
VINCULO_SEM_VINCULO = 'Sem vínculo familiar'

VINCULO_CHOICES = [
    (VINCULO_CONJUGE, 'Cônjuge'),
    (VINCULO_FILHO_ATE_6, 'Filho até 6 anos'),
    (VINCULO_FILHO_7_A_10, 'Filho de 7 a 10 anos'),
    (VINCULO_FILHO_11_A_23, 'Filho de 11 a 23 anos'),
    (VINCULO_FILHO_ACIMA_23, 'Filho acima de 23 anos'),
    (VINCULO_SEM_VINCULO, 'Sem vínculo familiar'),
]

# Quantidade de Hóspedes
QTDE_HOSP_CHOICES = [
    (1, '1'),
    (2, '2'),
    (3, '3'),
    (4, '4'),
    (5, '5'),
    (6, '6'),
]

def _build_uh_choices(total):
    return [(str(i), str(i)) for i in range(1, total + 1)]

# UH (Unidade Habitacional)
# HTM_01 tem 12 quartos; HTM_02 tem 4 quartos.
UH_CHOICES = _build_uh_choices(12)

# Especial
ESPECIAL_SIM = 'Sim'
ESPECIAL_NAO = 'Não'

ESPECIAL_CHOICES = [
    (ESPECIAL_SIM, 'Sim'),
    (ESPECIAL_NAO, 'Não'),
]

# MHEx
MHEX_HTM_01 = 'HTM_01'
MHEX_HTM_02 = 'HTM_02'

MHEx_CHOICES = [
    (MHEX_HTM_01, 'HTM_01'),
    (MHEX_HTM_02, 'HTM_02'),
]

# Mapeamento de UHs por hotel
UH_CHOICES_BY_MHEX = {
    MHEX_HTM_01: UH_CHOICES,
    MHEX_HTM_02: _build_uh_choices(3),
}

def get_uh_choices(mhex):
    return UH_CHOICES_BY_MHEX.get(mhex, UH_CHOICES)

# Status Reserva
STATUS_RESERVA_PENDENTE = 'Pendente'
STATUS_RESERVA_APROVADA = 'Aprovada'
STATUS_RESERVA_CHECKIN = 'Checkin'
STATUS_RESERVA_RECUSADA = 'Recusada'
STATUS_RESERVA_PAGO = 'Pago'
STATUS_RESERVA_EXPIRADA = 'Expirada'

STATUS_RESERVA_CHOICES = [
    (STATUS_RESERVA_PENDENTE, 'Pendente'),
    (STATUS_RESERVA_APROVADA, 'Aprovada'),
    (STATUS_RESERVA_CHECKIN, 'Checkin'),
    (STATUS_RESERVA_RECUSADA, 'Recusada'),
    (STATUS_RESERVA_PAGO, 'Pago'),
    (STATUS_RESERVA_EXPIRADA, 'Expirada'),
]

# Motivo Viagem
MOTIVO_SAUDE = 'Saúde'
MOTIVO_TRABALHO = 'Trabalho'
MOTIVO_TURISMO = 'Turismo'

MOTIVO_VIAGEM_CHOICES = [
    (MOTIVO_SAUDE, 'Saúde'),
    (MOTIVO_TRABALHO, 'Trabalho'),
    (MOTIVO_TURISMO, 'Turísmo'),
]

# UF
UF_CHOICES = [
    ('AL', 'AL'), ('AM', 'AM'), ('AP', 'AP'), ('BA', 'BA'), ('CE', 'CE'),
    ('DF', 'DF'), ('ES', 'ES'), ('GO', 'GO'), ('MA', 'MA'), ('MG', 'MG'),
    ('MS', 'MS'), ('MT', 'MT'), ('PA', 'PA'), ('PB', 'PB'), ('PE', 'PE'),
    ('PI', 'PI'), ('PR', 'PR'), ('RJ', 'RJ'), ('RN', 'RN'), ('RO', 'RO'),
    ('RR', 'RR'), ('RS', 'RS'), ('SC', 'SC'), ('SE', 'SE'), ('SP', 'SP'),
    ('TO', 'TO'),
]

# Produtos
PRODUTO_AGUA = 'Água'
PRODUTO_REFRIGERANTE = 'Refrigerante'
PRODUTO_CERVEJA = 'Cerveja'
PRODUTO_PET = 'Pet'
