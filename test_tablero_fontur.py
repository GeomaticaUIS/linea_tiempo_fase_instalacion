from datetime import datetime

from tablero_fontur import embarcadero_instalado


def test_embarcadero_instalado():
    assert embarcadero_instalado(datetime(2025, 7, 25)) is True
    assert embarcadero_instalado(None) is False
    assert embarcadero_instalado('') is False
