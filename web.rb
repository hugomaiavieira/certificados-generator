# encoding: utf-8
#
#  The MIT License
#
#  Copyright (c) 2011 Hugo Henriques Maia Vieira <hugomaiavieira@gmail.com>
#
#  Permission is hereby granted, free of charge, to any person obtaining a copy
#  of this software and associated documentation files (the "Software"), to
#  deal in the Software without restriction, including without limitation the
#  rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
#  sell copies of the Software, and to permit persons to whom the Software is
#  furnished to do so, subject to the following conditions:
#
#  The above copyright notice and this permission notice shall be included in
#  all copies or substantial portions of the Software.
#
#  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
#  FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
#  IN THE SOFTWARE.
#
#
# Version 1.0: 03/11/2011
# Version 2.0: 03/02/2012 - web user interface

require 'sinatra'
require 'odf-report'
require 'spreadsheet'
require 'ods'

UPLOAD_PATH = 'public/uploads'

get '/' do
  @errors = []
  erb :index
end

get '/ajuda' do
  erb :ajuda
end

post '/' do
  @errors = []
  validate(params)
  if @errors.empty?
    vars = params[:variaveis].split
    name = params[:nome].empty? ? 'certificados' : params[:nome]
    model = upload(params[:modelo], 'modelo.odt')

    data_mime = params[:dados][:type] =~ /ms-excel/ ? 'xls' : 'ods'
    data = upload(params[:dados], "dados.#{data_mime}")

    file = generate(vars, model, data)
    send_file file, :filename => "#{name}.odt"
  else
    erb :index
  end
end

get '/modelo-exemplo.odt' do
  send_file 'public/exemplos/modelo.odt'
end

get '/dados-exemplo.xls' do
  send_file 'public/exemplos/dados.xls'
end

def validate params
  if params[:modelo]
    @errors << 'O item Modelo tem que ser um arquivo odt' if not params[:modelo][:type] =~ /opendocument.text/
  else
    @errors << 'O item Modelo não pode ser vazio'
  end

  if params[:dados]
    @errors << 'O item Dados tem que ser um arquivo xls ou ods' if not params[:dados][:type] =~ /ms-excel|opendocument.spreadsheet/
  else
    @errors << 'O item Dados não pode ser vazio'
  end

  @errors << 'O item Variáveis não pode ser vazio' if params[:variaveis] and params[:variaveis].empty?
end

def upload multipart, filename
  file_path = File.join(UPLOAD_PATH, filename)

  File.open(file_path, "w") do |f|
    f.write(multipart[:tempfile].read)
  end

  return file_path
end

def generate(vars, model, data)
  file = 'public/output/certificados.odt'

  Spreadsheet.client_encoding = 'UTF-8'

  rows = data_rows(data)

  report = ODFReport::Report.new(model) do |r|
    r.add_section('CERTIFICADOS', rows) do |s|
      vars.each_with_index do |field, i|
        s.add_field(field.upcase) { |row| "#{row[i]}" }
      end
    end
  end

  report.generate(file)
  return file
end

def data_rows(data_path)
  puts data_path
  if data_path.end_with? '.xls'
    Spreadsheet.client_encoding = 'UTF-8'
    table = Spreadsheet.open(data_path).worksheets[0]
    rows = []; table.each { |row| rows << row }
    return rows
  else
    ods = Ods.new(data_path)
    sheet = ods.sheets[0]
    _rows = []
    sheet.rows.each do |row|
      _row = []
      row.cols.each { |cell| _row << cell.value }
      _rows << _row
    end
    return _rows
  end
end