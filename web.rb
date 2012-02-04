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

UPLOAD_PATH = 'public/uploads'

get '/' do
  erb :index
end

get '/ajuda' do
  erb :ajuda
end

post '/' do
  vars = params[:variaveis].split
  name = params[:nome].empty? ? 'certificados' : params[:nome]
  model = upload(params[:modelo], 'modelo.odt')
  data = upload(params[:dados], 'dados.xls')

  file = generate(vars, model, data)
  send_file file, :filename => "#{name}.odt"
end

get '/modelo-exemplo.odt' do
  send_file 'public/exemplos/modelo.odt'
end

get '/dados-exemplo.xls' do
  send_file 'public/exemplos/dados.xls'
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

  table = Spreadsheet.open(data).worksheets[0]
  rows = []; table.each { |row| rows << row }

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