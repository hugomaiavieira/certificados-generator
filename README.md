# Certificado Generator

Script para gerar certificados baseados em um modelo odt (LibreOffice Writer)
pegando os dados de um arquivo xls (Excel).

O arquivo modelo.odt deve ter uma seção¹ chamada CERIFICADOS. Dentro desta
seção, devem ser inseridas as variáveis - por exemplo [PALESTRANTE], [PALESTRA],
[DIA] e [ETC] - que serão substituídas pelos valores contidos no arquivo
dados.xls. A ordem dos parâmetros na chamada do programa deve seguir a ordem das
colunas do arquivo dados.xls.

O nome da seção e das variáveis devem estar em letras MAIÚSCULAS, sendo que as
variáveis devem estar ente [COLCHETES].

    Exemplo de uso:

        $ ./certificados-generator palestrante palestra dia

¹ Para adicionar uma seção, vá no menu Inserir -> Seção.

# Licença

The MIT License

Copyright (c) 2011 Hugo Henriques Maia Vieira

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to
deal in the Software without restriction, including without limitation the
rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
IN THE SOFTWARE.

