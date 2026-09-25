from datetime import datetime

from tablero_fontur import embarcadero_instalado, parsear_celda


def test_embarcadero_instalado():
    assert embarcadero_instalado(datetime(2025, 7, 25)) is True
    assert embarcadero_instalado(None) is False
    assert embarcadero_instalado('') is False


def test_parsear_celda_no_requiere_incluye_archivo():
    resultado = parsear_celda('NO REQUIERE')
    assert resultado['estado'] == 'no_requiere'
    assert resultado['ruta'] == ''
    assert resultado['archivos'] == ['NO REQUIERE']


def test_parsear_celda_con_ruta_y_archivo():
    valor = 'Documentos\\ruta\\archivo.pdf\nNOMBRE ARCHIVO: archivo.pdf'
    resultado = parsear_celda(valor)
    assert resultado['estado'] == 'con_soporte'
    assert resultado['ruta'].endswith('ruta') or resultado['ruta'] == 'Documentos\\ruta\\archivo.pdf'
    assert resultado['archivos'] == ['archivo.pdf']
