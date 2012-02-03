require 'sinatra'

get '/' do
  erb :index
end

get '/modelo-exemplo.odt' do
  send_file 'public/exemplos/modelo.odt'
end

get '/dados-exemplo.xls' do
  send_file 'public/exemplos/dados.xls'
end