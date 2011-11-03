#!/usr/bin/env ruby
# encoding: utf-8
#
#  Script para gerar certificados baseados em um modelo odt (LibreOffice Writer)
#  pegando os dados de um arquivo xls (Excel).
#
#  O arquivo modelo.odt deve ter uma seção¹ chamada CERIFICADOS. Dentro desta
#  seção, devem ser inseridas as variáveis - por exemplo [PALESTRANTE],
#  [PALESTRA], [DIA] - que serão substituídas pelos valores contidos no arquivo
#  dados.xls. A ordem de chamada do programa deve seguir a ordem das colunas do
#  arquivo dados.xls.
#
#  O nome da seção e das variáveis devem estar em letras MAIÚSCULAS, sendo que
#  as variáveis devem estar ente [COUCHETES].
#
#  Exemplo de uso:
#
#      $ ./certificados-generator palestrante palestra dia
#
#  ¹ Para adicionar uma seção, vá no menu Inserir -> Seção.
#
# Autor:  Hugo Maia Vieira <hugo@algorich.com.br>
# Data:   03/11/2011

require "rubygems"
require 'odf-report'
require 'spreadsheet'

PATH = File.expand_path(File.dirname(__FILE__))
DATA = File.join(PATH, 'dados.xls')
MODEL = File.join(PATH, 'modelo.odt')

Spreadsheet.client_encoding = 'UTF-8'

table = Spreadsheet.open(DATA).worksheets[0]
rows = []; table.each { |row| rows << row }

report = ODFReport::Report.new(MODEL) do |r|
  r.add_section('CERTIFICADOS', rows) do |s|
    ARGV.each_with_index do |field, i|
      s.add_field(field.upcase) { |row| "#{row[i]}" }
    end
  end
end

report.generate("./arquivos/certificados.odt")

