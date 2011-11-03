#!/usr/bin/env ruby
# encoding: utf-8
#
#  Script para gerar certificados baseados em um modelo em odt.
#
#  O arquivo modelo.odt deve ter uma seção¹ chamada CERIFICADOS. Dentro desta
#  seção, devem existir as variáveis [PALESTRANTE], [PALESTRA] e [DIA], que
#  serão substituidas pelos valores contidos no arquivo do Excel.
#
#  O arquivo quem contém os nomes deve se chamar dados.xls. Todas as informações
#  de cada palestrante devem estar em uma linha. As colunas devem estar
#  exatamente na ordem: nome do palestrante, titulo da palestra, dia da palestra
#
#  ¹ Para adicionar uma seção, vá no menu Inserir -> Seção.
#
# Autor:  Hugo Maia Vieira <hugo@algorich.com.br>
# Data:   03/11/2011

require "rubygems"
require 'odf-report'
require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'

tabela = Spreadsheet.open('./dados.xls').worksheets[0]

linhas = []; tabela.each { |linha| linhas << linha }

report = ODFReport::Report.new("./modelo.odt") do |r|
  r.add_section("CERTIFICADOS", linhas) do |s|
    s.add_field('PALESTRANTE') { |linha| "#{linha[0]}" }
    s.add_field('PALESTRA') { |linha| "#{linha[1]}" }
    s.add_field('DIA') { |linha| "#{linha[2]}" }
  end
end

report.generate("./arquivos/certificados.odt")

