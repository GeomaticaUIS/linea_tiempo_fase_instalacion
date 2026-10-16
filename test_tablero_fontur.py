from datetime import date, datetime

from tablero_fontur import embarcadero_instalado, formatear_fecha_instalacion, parsear_celda


def test_embarcadero_instalado():
    assert embarcadero_instalado(datetime(2025, 7, 25)) is True
    assert embarcadero_instalado(None) is False
    assert embarcadero_instalado('') is False


def test_formatear_fecha_instalacion():
    assert formatear_fecha_instalacion(datetime(2025, 7, 25)) == '25/07/2025'
    assert formatear_fecha_instalacion(date(2025, 7, 25)) == '25/07/2025'
    assert formatear_fecha_instalacion(None) == ''


def test_parsear_celda_no_requiere_incluye_archivo():
    resultado = parsear_celda('NO REQUIERE')
    assert resultado['estado'] == 'no_requiere'
    assert resultado['ruta'] == ''
    assert resultado['archivos'] == ['NO REQUIERE']


def test_parsear_celda_no_aplica_no_cuenta_como_pendiente():
    resultado = parsear_celda('NO APLICA')
    assert resultado['estado'] == 'no_requiere'
    assert resultado['ruta'] == ''
    assert resultado['archivos'] == ['NO APLICA']


def test_parsear_celda_instalado_por_decreto_se_trata_como_no_requiere():
    for mensaje in (
        'INSTALADO POR DECRETO DE CALAMIDAD PÚBLICA',
        'INSTALADO POR DECRETO DE CALAMIDAD P\ufffdBLICA',
    ):
        resultado = parsear_celda(mensaje)
        assert resultado['estado'] == 'no_requiere'
        assert resultado['ruta'] == ''
        assert resultado['archivos'] == ['INSTALADO POR DECRETO DE CALAMIDAD PÚBLICA']


def test_parsear_celda_con_ruta_y_archivo():
    valor = 'Documentos\\ruta\\archivo.pdf\nNOMBRE ARCHIVO: archivo.pdf'
    resultado = parsear_celda(valor)
    assert resultado['estado'] == 'con_soporte'
    assert resultado['ruta'].endswith('ruta') or resultado['ruta'] == 'Documentos\\ruta\\archivo.pdf'
    assert resultado['archivos'] == ['archivo.pdf']


def test_parsear_celda_con_hyperlink():
    valor = 'Documentos\\Documentos Tecnicos\\ruta\nNOMBRE ARCHIVO: 01_BTC_PtoWilches_140725'
    resultado = parsear_celda(valor, href='https://sharepoint.example/bitacora')
    assert resultado['estado'] == 'con_soporte'
    assert resultado['href'] == 'https://sharepoint.example/bitacora'
