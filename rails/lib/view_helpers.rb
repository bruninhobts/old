require 'net/http'
require 'uri'

module PagSeguroViewHelper
  
  include Pagseguro
  
  # Cria um form para a finalizacao da compra - a tela final de uma compra.
  # 
  # O metodo retira os produtos da sessao (:carrinho)
  #
  # Voce pode passar como argumento um <tt>:mostrar_dados_cliente => true</tt> para que seja mostrado
  # os campos com os dados do cliente (esses dados serao enviados para o PagSeguro e esse nao ira solicita-los
  # de novo).
  def mostra_carrinho(*args)
    
    carrinho = session[:carrinho]
    
    result =  "<form target='pagseguro' method='post' action='https://pagseguro.uol.com.br/security/webpagamentos/webpagto.aspx'>\n"
    result << "<input type='hidden' name='email_cobranca' value='#{email}' />\n"
    result << "<input type='hidden' name='tipo' value='CP'/>\n"
    result << "<input type='hidden' name='moeda' value='BRL' />\n"
    
    if carrinho
      carrinho.values.each_with_index {|produto_quantidade, i|
        i += 1
        produto = produto_quantidade[:produto]
        result << "<input type='hidden' name='item_id_#{i}' value='#{produto.id}' />\n"
        result << "<input type='hidden' name='item_descr_#{i}' value='#{produto.descricao}' />\n"
        result << "<input type='hidden' name='item_valor_#{i}' value='#{produto.valor}' />\n"
        result << "<input type='hidden' name='item_quant_#{i}' value='#{produto_quantidade[:quantidade]}' />\n"
      }
    end
    
    if args and args.first.instance_of? Hash and args.first.keys.first == :mostrar_dados_cliente
      result << '''
      <label>Nome </label><input type="input" name="cliente_nome" value="" /><br />
      <label>CEP </label><input type="input" name="cliente_cep" value="" /><br />
      <label>Endereco </label><input type="input" name="cliente_end" value="" /><br />
      <label>Numero </label><input type="input" name="cliente_num" value="" /><br />
      <label>Complemento </label><input type="input" name="cliente_compl" value="" /><br />
      <label>Bairro</label><input type="input" name="cliente_bairro" value="" /><br />
      <label>Cidade </label><input type="input" name="cliente_cidade"value="" /><br />
      <label>UF </label><input type="input" name="cliente_uf" value="" /><br />
      <label>Pais </label><input type="input" name="cliente_pais" value="" /><br />
      <label>DDD </label><input type="input" name="cliente_ddd" value="" /><br />
      <label>Telefone </label><input type="input" name="cliente_tel" value="" /><br />
      <label>Email </label><input type="input" name="cliente_email" value="" /><br />
      '''
    end
    
    result << """<input type='image' src='https://pagseguro.uol.com.br/Security/Imagens/btnfinalizaBR.jpg' 
                        name='submit' alt='Pague com PagSeguro - é rápido, grátis e seguro!' />"""
                        
    result << "</form>"
  end
  
  # Cria um form hidden para compra de apenas um produto
  #  
  # O parametro (produto) deve ser um objeto que responda ao seguintes metodos:
  #  * id (para identificar o produto)
  #  * descricao
  #  * valor
  #  
  # Opcionalmente o objeto tambem pode responder a frete - nesse caso
  # um campo hidden adicional sera colocado.
  def form_para_compra(produto)
    
    validar_produto produto
    
    result =  "<form target='pagseguro' method='post' action='https://pagseguro.uol.com.br/security/webpagamentos/webpagto.aspx'>"
    result << "<input type='hidden' name='email_cobranca' value='#{email}' />\n"
    
    result << "<input type='hidden' name='tipo' value='CP'/>\n"
    result << "<input type='hidden' name='moeda' value='BRL' />\n"
    
    result << "<input type='hidden' name='item_id_1' value='#{produto.id}' />\n"
    result << "<input type='hidden' name='item_descr_1' value='#{produto.descricao}' />\n"
    result << "<input type='hidden' name='item_valor_1' value='#{produto.valor}' />\n"
    
    if produto.respond_to? :frete
      result << "<input type='hidden' name='frete value='#{produto.frete}' /> \n"
    else
      result << "<input type='hidden' name='frete' value='0' />\n"
    end
    
    result << "<input type='hidden' name='item_quant_1' value='1' />\n"
    result << "<input type='hidden' name='peso' value='100' />\n"
    
    result << """<input type='image' src='https://pagseguro.uol.com.br/Security/Imagens/btnfinalizaBR.jpg' 
                        name='submit' alt='Pague com PagSeguro - é rápido, grátis e seguro!' />"""
    result << "</form>"
  end
  
  
  # Adiciona um produto no carrinho - o carrinho e um hash aonde o id do produto e o valor 
  # e um hash em que: 
  #
  # <tt>{ :produto => produto_em_si, :quantidade => quantidade_de_produtos_no_carrinho }</tt>
  def adiciona_produto_carrinho(produto)
    
    validar_produto produto
    
    carrinho = session[:carrinho]
    unless carrinho
      carrinho = {}
    end
    
    if carrinho[produto.id]
      carrinho[produto.id][:quantidade] += 1
    else
      carrinho[produto.id] = {:produto => produto, :quantidade => 1}
    end
    
    session[:carrinho] = carrinho
  end
  
  # Retorna um array de produtos que estao no carrinho no momento (sem a informacao de quantidade)
  def produtos_no_carrinho
    if session[:carrinho]
      session[:carrinho].values.collect { |info| info[:produto] }
    else
      []
    end
  end
  
  # Remove um produto do carrinho pelo id do mesmo, caso haja mais de um produto, em quantidade,
  # decresce a quantidade - se nao, o remove completamente.
  def remover_produto_carrinho(id)
    carrinho = session[:carrinho]
    produto = carrinho[id]
    
    if produto[:quantidade] > 1
      produto[:quantidade] -= 1
      carrinho[id] = produto
    else
      carrinho.delete id
    end
    
    session[:carrinho] = carrinho
    
  end
  
  # Processa o retorno do pagseguro - verifique isso funciona em http://visie.com.br/pagseguro/retorno-automatico.php
  def processa_retorno(params)
    
    unless validar_pedido
      return nil
    end
    
    retorno = {}
    
    numItens = params[:NumItens].to_i
    
    produtos = []
    
    retorno[:transacao_id] = params[:TransacaoID]
    retorno[:tipo_frete] = params[:TipoFrete]
    retorno[:valor_frete] = params[:ValorFrete]
    retorno[:anotacao] = params[:Anotacao]
    retorno[:data_transacao] = DateTime.parse(params[:DataTransacao])
    retorno[:tipo_pagamento] = params[:TipoPagamento]
    retorno[:status_transacao] = params[:StatusTransacao]
    
    retorno[:cliente] = { :nome => params[:CliNome], :email => params[:CliEmail], :endereco => params[:CliEndereco],
                          :numero => params[:CliNumero], :complemento => params[:CliComplemento], :bairro => params[:CliBairro],
                          :cidade => params[:CliCidade], estado => params[:CliEstado], :cep => params[:CliCEP], 
                          :telefone => params[:CliTelefone]
                        }
    
    (1..numItens).each {|num|
      produtos << { :id => params[:"ProdID_#{num}"], :descricao => params[:"ProdDescricao_#{num}"], :valor => params[:"ProdValor_#{num}"], 
                    :quantidade => params[:"ProdQuantidade_#{num}"], :frete => params[:"ProdFrete_#{num}"], 
                    :taxas => params[:"ProdExtras_#{num}"]
                  }

    }
    
    retorno["produtos"] = produtos
    
    retorno
  end
  
  # Faz um post para a url do pagseguro para descobrir se um pedido e valido.
  def validar_pedido 
    res = Net::HTTP.post_form(URI.parse('https://pagseguro.uol.com.br/Security/NPI/Default.aspx/'), 
                              {'Comando'=>'validar', 'token'=> token })
                              
    res == "VERIFICADO"
  end

  private 

  # Valida se um objeto responde aos metodos requiridos pelo plugin
  def validar_produto(produto)
    unless produto.respond_to? :id and produto.respond_to? :descricao and produto.respond_to? :valor
      throw ArgumentError.new("Seu produto deve responder aos seguintes metodos: :id, :valor e :descricao")
    end
  end
end