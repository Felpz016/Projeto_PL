"""Busca de preços ao vivo em APIs públicas (sem chave)."""
import requests

AZURE_URL = "https://prices.azure.com/api/retail/prices"


def azure_preco_hora(sku: str, regiao: str) -> float:
    """Preço pay-as-you-go em US$/hora de uma VM Linux na Azure Retail Prices API."""
    filtro = (f"armRegionName eq '{regiao}' and armSkuName eq '{sku}' "
              "and priceType eq 'Consumption' and serviceName eq 'Virtual Machines'")
    r = requests.get(AZURE_URL, params={"$filter": filtro}, timeout=15)
    r.raise_for_status()
    itens = [i for i in r.json().get("Items", [])
             if "Windows" not in i.get("productName", "")
             and "Spot" not in i.get("skuName", "") and "Low Priority" not in i.get("skuName", "")]
    if not itens:
        raise ValueError("SKU/região não encontrado")
    return float(min(i["retailPrice"] for i in itens))
