# Certificados Generator

Script para gerar certificados baseados em um modelo odt ([LibreOffice][1]
Writer) pegando os dados de um arquivo xls (Excel).

## Instalação

Primeiro faça o [download aqui][2]. Em seguida, extraia a pasta contida no
arquivo `hugomaiavieira-certificados-generator-*` em um diretório qualquer e, se
quiser, renomeie a pasta extraída para `certificados-generator`. Por exemplo, eu
extrai em `/home/hugo/Documentos` e renomeie.

Em seguida, abra o Terminal e digite os comandos:

    $ sudo apt-get install ruby rubygems
    $ sudo apt-get install libxslt-dev libxml2-dev
    $ cd ~/Documentos/certificados-generator
    $ gem install bundler --no-rdoc --no-ri
    $ bundle install

## Utilização

1. Na pasta do programa, deve-se ter um aquivo chamado modelo.odt. Este arquivo
servirá como modelo para a geração dos certificados.
2. O arquivo modelo.odt deve ter uma seção¹ chamada CERIFICADOS.
3. Dentro desta seção, devem ser inseridas as variáveis - por exemplo
[PALESTRANTE], [PALESTRA], [DIA] e [ETC].
4. Na pasta do programa, deve-se ter um arquivo chamado dados.xls. Este arquivo
irá conter os dados que serão substituídos no lugar das variáveis.
3. O nome da seção e das variáveis devem estar em letras MAIÚSCULAS, sendo que
as variáveis devem estar ente [COLCHETES].
4. A ordem dos parâmetros na chamada do programa deve seguir a ordem das colunas
do arquivo dados.xls.

¹ Para adicionar uma seção, vá no menu Inserir -> Seção.

### Exemplo de uso

Supondo que eu meu arquivo dados.xls tenha os seguintes dados (veja o arquivo
dados_exemplo.xls):

<table>
    <tr>
        <td>Hugo Maia Vieira</td>
        <td>Como gerar certificados</td>
        <td>02/01/2012</td>
    </tr>
        <td>Thomas Edison</td>
        <td>A invenção da lâmpada</td>
        <td>14/08/1879</td>
    <tr>
    </tr>
    <tr>
        <td>Júlio Verne</td>
        <td>Viagem ao centro da terra</td>
        <td>21/05/1864</td>
    </tr>
</table>

Assim, vou criar o meu modelo.odt com as variáveis [PALESTRANTE], [PALESTRA] e
[DIA], lembrando que as variáveis devem estar dentro de uma seção¹ chamada
CERIFICADOS (veja o arquivo modelo_exemplo.odt).

Para abrir o terminal para executar o programa, vá na pasta onde foi este foi
extraído, clique duas vezes no arquivo `executar` e depois pressione "Executar".

O terminal irá abrir esperando que você chame o programa. A ordem dos parâmetros
na chamada do programa deve seguir a ordem das colunas do arquivo dados.xls.
Para nosso exemplo, vamos chamar o programa dessa maneira e apertar "Enter":

    $ ./certificados-generator palestrante palestra dia

Obs: não se esqueça do `./` no início da chamada

Pronto! Os certificados serão gerados em um único arquivo, dentro da pasta
arquivos. Agora basta imprimir.

## Licença

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

[1]: http://www.libreoffice.org/
[2]: https://github.com/hugomaiavieira/certificados-generator/tarball/master