#!/usr/bin/env ruby
# encoding: utf-8
#
#  The MIT License
#
#  Copyright (c) 2011 Hugo Henriques Maia Vieira
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
# Author: Hugo Maia Vieira <hugo@algorich.com.br>
#
# Version 1.0: 03/11/2011

require "rubygems"
require 'odf-report'
require 'spreadsheet'

PATH = File.expand_path(File.dirname(__FILE__))
DATA = File.join(PATH, 'dados.xls')
MODEL = File.join(PATH, 'modelo.odt')
FILE = File.join(PATH, 'arquivos/certificados.odt')

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

report.generate(FILE)

