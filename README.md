# format-barcode128
Converte uma string de caracteres em uma string 128B ou um 128C  de código de barras para uso em Excel ou Word
# Input
string ou variável contendo uma string para conversão
# Output
string de código de barras com checksum
# forma de uso
chamar a função Cb128(string ou varável)
# Uso básico em Excel
//seta a fonte Code 128 no range de células <br>
    chWorkSheet:Range("K1:K18"):FONT:NAME = "Code 128". <br>
//chama a função que fará a conversão da string <br>
    chWorksheet:range("K1:K18"):VALUE = Cb128(c-string). 
