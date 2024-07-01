# Lista de bancos (puedes agregar más bancos aquí)
def get_bancos():
    bancos = ["BBVA", "Santander", "Galicia"]

    return bancos


def get_codigos_bancos():
    codigos_bancos = {
        "BBVA": {
            # se usa com m para evitar que se incluyan conceptos como compras (por ejemplo)
            "similares": [
                ["com m", "comi", "comision"],
            ],
            "no_similares": ["sircreb", "25413", "iva tasa", "perc.caba", "percepcion iva"]
        },
        "Santander": {
            "similares": [],
            "no_similares": ["santander_code1", "santander_code2"]
        },
        "Galicia": {
            "similares": [],
            "no_similares": ["galicia_code1", "galicia_code2"]
        }
    }

    return codigos_bancos
