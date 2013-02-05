carrinho={}
for(i in produtos)carrinho[i]=0

function formata(n){
  n=String(n)
  if(n.indexOf(".")==-1)return n+",00"
  return (n+"00").replace(/\.(\d\d).*/,",$1")
}

function mostraCarrinho(){
  total=0
  t="<table><tr><th colspan='6'>Suas Compras</th></tr>"
  t+= "  <tr>";
  t+= "    <td width='30' align='center'><b>ID</b></td>";
  t+= "    <td width='200' align='center'><b>Produto</b></td>";
  t+= "    <td width='60' align='center'><b>Quant</b></td>";
  t+= "    <td width='60' align='center'><b>Valor</b></td>";
  t+= "    <td width='60' align='center'><b>Total</b></td>";
  t+= "    <td width='60' align='center'><b>Excluir</b></td>";
  t+= "  </tr>";
  for(i in carrinho){
    if(carrinho[i]){
      total+=Number(produtos[i][1].replace(/,/,"."))*carrinho[i]
      t+= "  <tr>";
      t+= "    <td>" + i + "</td>";
      t+= "    <td>" + produtos[i][0] + "</td>";
      t+= "    <td align='center'><input type='text' value='"+carrinho[i]+"' size='2' onblur='upP(\"" + i + "\",this.value)'></td>";
      t+= "    <td align='right'>" + produtos[i][1] + "</td>";
      t+= "    <td align='right'>" + 
          formata(Number(produtos[i][1].replace(/,/,"."))*carrinho[i]) + "</td>";
      t+= "    <td align='center'><a href='#' onclick='if(confirm(\"Excluir?\"))delP(\"" + i + "\");'>Excluir</a></td>";
      t+= "  </tr>";
    }
  }
  t+="</table><span><b>Total: "+formata(total)
  t+=' <img src="https://pagseguro.uol.com.br/Security/Imagens/btnfinalizaBR.jpg" '
  t+='onclick="comprar()" alt="Pague com PagSeguro - é rápido, grátis e seguro!" '+
     'style="cursor:pointer">'
  t+=" </b> </span>"
  $("#cart").html(total?t:"")
}

function addP(i){
  if(String(carrinho[i])=="NaN")carrinho[i]=0
  carrinho[i]+=1
  mostraCarrinho()
}

function upP(i,q){
  carrinho[i]=Number(q)
  mostraCarrinho()
}

function delP(i){upP(i,0)}

$(function(){
  t="<table>"
  for(i in produtos){
    t+="<tr><td>"+produtos[i][0]+"</td><td><input type='button'"+
       "value='Adicionar' onclick='addP(\""+i+"\")' /></td></tr>"
  }
  t+="</table>"
  $("#lista").html(t)
})

function hidinput(n,v){
  return '<input type="hidden" name="'+n+'" value="'+v+'">'
}

function comprar(){
  ff='<form target="pagseguro" method="post" action="https://pagseguro.uol.com.br/security/webpagamentos/webpagto.aspx">'
  ff+=hidinput('email_cobranca',email_cobranca)+
      hidinput('tipo','CP')+
      hidinput('moeda','BRL')+
      hidinput('tipo_frete','EN')
  c=0
  for(i in carrinho){
    if(carrinho[i]){
      c++
      ff+=hidinput('item_id_'+c,i)+
          hidinput('item_descr_'+c,produtos[i][0])+
          hidinput('item_quant_'+c,carrinho[i])+
          hidinput('item_valor_'+c,produtos[i][1].replace(/,/,''))+
          hidinput('item_peso_'+c,0)
    }
  }
  ff+='</form>'

  $("body").append(ff)
  $("form:last").submit()
}
