from datetime import datetime

def ordenar_por_validade(lista, crescente=True):
    def parse_data(prod):
        try:
            data = prod['validade'] if isinstance(prod, dict) else prod[3]
            return datetime.strptime(data, "%d/%m/%Y")
        except:
            return datetime.max if crescente else datetime.min
    return sorted(lista, key=parse_data, reverse=not crescente)
