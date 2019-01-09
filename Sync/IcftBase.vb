' --------------------------------------------------------------------------------
'
' INTERCRAFT SOLUTIONS INFORMÁTICA LTDA
' 14 DE JUNHO DE 2007 - BASE COMUM PARA SOLUÇÕES ICRAFT
' BIBLIOTECA PADRÃO PARA SITE ASPNET E APL VB
'
' OBSERVAÇÕES:::
'   - xxobservaçãoxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'   - qualquer alteração deverá ser enviada por msg a toda equipe desenv@icraft.com.br.
'   - funções db devem considerar     oracle/mysql/msaccess.
'   - tipos equivalentes em textlong  clob/longtext/memo.
'   - tipos equivalentes em binary    blob/blob/olebinary.
'   - webconn considera para site asp.net o web.config e para apl windows app.config.
'   - todas as funções classes enumerações precisam de esclarecimentos --> '''.
'
' ALTERAÇÕES:::
'   - dd/mmm/yyyy xtécnico xxmotivotodocomletrasminúsculasxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'   - 14/jul/2007 lucianol normalização, testes clob/blob e inclusão de objecttobytearray.
'   - 14/jul/2007 lucianol acerto nz para considerar nulo quando isnothing.
'   - 14/jul/2007 lucianol inclusão função dsproxseq para retornar próximo sequencial (simula sequence).
'   - 17/jul/2007 lucianol inclusão função dscarregacampos para preenchimento automático de campos de formulário.
'   - 18/jul/2007 lucianol inclusão função dsgravacampos para gravação de campos automaticamente em tabela ou sql.
'   - 21/jul/2007 lucianol gravação binária testes com oracle mysql access criação função ByteArrayToObject e CampoConteudo.
'   - 22/jul/2007 lucianol tratamento de erro em conversões bytearrayobject e vice-versa retornando Nothing quando um erro ocorrer.
'   - 23/jul/2007 lucianol transformação função controleconteudo para classe.
'   - 30/jul/2007 lucianol criação dscarregatop.
'   - 01/ago/2007 thiagop  inclusão de função ShowJSMessage para exibir mensagem de aviso após submissão de dados.
'   - 13/set/2007 lucianol separação de funções obsoletas, preparação da prop para procurar controles em containers.
'   - 03/out/2007 thiagop  inclusão de funções ExibeData, GravaData e Enum para Parâmetro na GravaData.
'   - 06/nov/2007 lucianol findcontrol normal não encontrava itens em paineis existentes (panel.findcontrol). inclui findcontrolespecial.
'   - 07/nov/2007 lucianol criação de componentes icraftcombobox e icraftgridview.
'   - 08/nov/2007 lucianol criação de rotinas de combobox tratamento padrão addhandle para mudança de itens dependentes.
'   - 08/nov/2007 lucianol alteração de rotina de busca de componentes automática para form.findgeral, que procura nos filhos e nos pais.
'   - 08/nov/2007 lucianol agrupamento e identificação de rotinas combobox separando-as no código para melhor visualização. criação de constantes combonull e combosepdefault.
'   - 08/nov/2007 lucianol exclusão da função dscarregacombo, substituída por carregacombo em seu novo formato.
'   - 08/nov/2007 lucianol origemcontrole, rotina excluída, substituída por form.findgeral...
'   - 08/nov/2007 lucianol icftcombobox, inclusão de opção de lista no combobox.
'   - 08/nov/2007 lucianol macrosubstsql, possibilidade de fazer referência ao valor diretamente. Definição explícita de tipos param esperados.
'   - 08/nov/2007 lucianol macrosubstsql, previsão de tratamento do tipo [:exprsql.FLTCAMPO], que é igual a [:FLTCAMPO].
'   - 08/nov/2007 lucianol macrosubstsql, previsão de tratamento do tipo [:valor.FLTCAMPO] para substituição pura e simples no sql sem considerar EXPRSQL. Utilizado para colocar nomes de campos etc.
'   - 09/nov/2007 lucianol macrosubsttext, sendo função prevendo apenas o retorno de sql com traduções.
'   - 09/nov/2007 lucianol nzv, considerando zero também como valor vazio de numéricos (como "" para string).
'   - 09/nov/2007 lucianol regexmatches, retirada do parâmetro group, já que não estava sendo utilizado.
'   - 10/nov/2007 lucianol html.a_ref, colocação do httputility.htmlencode para qualquer código ser interpretado como html.
'   - 10/nov/2007 lucianol erromsg, inserção de rotina de notificação e tratamento de erro para evitar uso da tela showjsmessage.
'   - 10/nov/2007 lucianol geradefs, rotinas para leitura dos atributos das tabelas do gerador (sistema, tabela, campos etc.).
'   - 11/nov/2007 lucianol icftcombobox, criação de propriedades style e containerstyle, impl da rotina de erro, testes de funcionamento.
'   - 11/nov/2007 lucianol icftcombobox, ajuste de rotina de tramento de erro e retirada dos try catchs das funções da biblioteca.
'   - 11/nov/2007 lucianol icftdetalhes, inclusão de controle para detalhes de registro.
'   - 11/nov/2007 lucianol nz, previsão de tipo int32, que faltava.
'   - 13/nov/2007 lucianol notamsg, troca do nome da classe de erroicft para notaicft, para permitir qq tipo de notificação.
'   - 13/nov/2007 lucianol html, inclusão de código para criação de tabelas.
'   - 13/nov/2007 lucianol gerador, classe sendo adaptada para importar e exportar formatos de estrutura.
'   - 19/nov/2007 lucianol extendtoarray, para permitir que um campo guarde inúmeros valores.
'   - 19/nov/2007 lucianol temporaryfile, para retornar nome de arquivo livre para tratamento (deve ser excluído no final do procedimento).
'   - 19/nov/2007 lucianol temporarydir, obtém diretório temporário, que corresponde ao param de config dir_temp.
'   - 19/nov/2007 lucianol gerador, recurso para exportar em oracle, criando triggers e cascate update (sem recursividade).
'   - 19/nov/2007 lucianol dscarregacombo, retirada do recurso e verificação de todo o código.
'   - 20/nov/2007 lucianol extendtoarraylist, passa do tipo armazenado em campo texto para arraylist.
'   - 20/nov/2007 lucianol procuranode, procura node em uma árvore através de um de seus atributos.
'   - 20/nov/2007 lucianol inseretab, inclui no bloco de texto tab antes de cada linha, considerando separador como parâmetro.
'   - 20/nov/2007 lucianol filtroform, diferenciação de texto para apresentação do filtro simplificado em tela para usuario.
'   - 20/nov/2007 lucianol filtroform, atualização de parâmetros de controle de filtro para mostrar etiqueta e tooltip no form.
'   - 13/dez/2007 lucianol downloadarquivo, para facilitar o processo de envio de arquivo para browser do usuário.
'   - 13/dez/2007 lucianol downloadconteudo, para envio de conteúdo ao cliente.
'   - 13/dez/2007 lucianol prop, alteração para buscar o objeto, quando este for passado por texto ou retornar nothing.
'   - 13/dez/2007 lucianol gerador, função diferença percebendo mais parâmetros além de nome etc entre oracle e definições.
'   - 13/dez/2007 lucianol gerador, exportação atualizada da estrutura para oracle e access.
'   - 13/dez/2007 lucianol prop, previsão da propriedade checked para combobox put e set.
'   - 13/dez/2007 lucianol gerador, inclusão de rotina para documentar diferença entre relacionamentos.
'   - 13/dez/2007 lucianol gerador, rotina de convesão de tipos oracle para access.
'   - 20/dez/2007 lucianol autoseq, inclusão de tratamento de campos automáticos. mencionando proxseq no gerador, campo é atualizado com próximo sequencial automaticamente.
'   - 20/dez/2007 lucianol valorpadrao, inclusão de tratamento do campo gerador valorpadrao, permitindo NOW para formatação de data atual.
'   - 20/dez/2007 lucianol filtro em form, inclusão de campo genérico para filtro de todos os campos.
'   - 23/dez/2007 lucianol segmexpr, inclusão da função para concatenar qualquer lista de segmentos considerando separador.
'   - 23/dez/2007 lucianol controle, preparação para considerar tipo de campo na formatação de entrada e apresentação.
'   - 23/dez/2007 lucianol formatos possíveis outros: | memo | html.
'   - 23/dez/2007 lucianol formatos possíveis data: | dd/MM/yyyy | dd/MM/yyyy HH:mm | dd/MM/yyyy HH:mm:ss.
'   - 23/dez/2007 lucianol formatos possíveis número: | inteiro | real.
'   - 25/dez/2007 lucianol regex mudança de parâmetro de grupo para permitir nomes em regex ?<nome>...
'   - 27/dez/2007 lucianol comandoaccess alteração da rotina para converter código de concatenação " || " em " & " (observar os espaços).
'   - 05/jan/2008 lucianol criacampo, inclusão desta rotina na biblioteca para criação dinâmica de campos.
'   - 05/jan/2008 lucianol rel1n, inclusão desta classe na biblioteca para criação de combobox dinâmica baseada em relacionamentos.
'   - 05/jan/2008 lucianol troca de dd/MM/yyyy para dd\/MM\/yyyy para manter a barra mesmo com idioma inglês.
'   - 05/jan/2008 lucianol mudei atualizoucombo para atualizoucontrole, pois função serve para todos os controles.
'   - 05/jan/2008 lucianol erro na função atualiza combo com mais de uma coluna. estava desconsiderando ocultavalor. erro de escrita da do attributo foi corrigido.
'   - 05/jan/2008 lucianol crypb decrypb inclusão de rotinas de criptografia básica, só para esconder código do usuário em atributo de campo.
'   - 05/jan/2008 lucianol icftdetails criação de componente para permitir alteração de registros baseando-se em sql simples, orientado pela chave.
'   - 05/jan/2008 lucianol icftdetails vínculo entre details e gridview, permitindo edição simplificada de registros de cadastro.
'   - 05/jan/2008 lucianol icftform correção de aplicação de filtro vazio, que apresentava erro.
'   - 05/jan/2008 lucianol rotinas de linha, inclusão das rotinas de tratamento de buffer de linha (de registro) na biblioteca funções para salvar, checar e comparar.
'   - 10/jan/2008 lucianol nz, tratamento de nulo em conteúdos numéricos, que geravam erro. nulo retornará zero.
'   - 10/jan/2008 lucianol idioma, registro de idioma em variável de sessão com verificação de target e argumentos para mudança.
'   - 14/jan/2008 thiagop  função IncluiStyleSheet, para inclusão dinâmica de folhas de estilo na página.
'   - 21/fev/2008 lucianol automatizagrid, que inclui diversas funcionalidades em grid simples, conforme padronização
'   - 01/mar/2008 lucianol atualização da biblioteca icraft.vb com suporte aplicativo form além da web.
'   - 11/jul/2008 lucianol dsgrava e similares para mysql utilizar registro de variáveis de ambiente: CONN_IP, CONN_MACHINE E CONN_USER.
'   - 11/jul/2008 lucianol gerador mysql inclusão de logon indireto fazendo uso das variáveis CONN_IP, CONN_MACHINE E CONN_USER.
'   - 27/nov/2008 lucianol --- remodelagem de código com base nas necessidades evitando dependência de objeto DAO, sendo este criado conforme seu uso.
'   - 27/nov/2008 lucianol utilização de begin area para organização de código.
'   - 28/nov/2008 lucianol explicações em todas as funções e trechos do código (que estavam faltando).
'   - 28/nov/2008 lucianol correção em rotina pausa, que na Internet não aguardava segundos, System.Threading.Thread.Sleep(Segundos * 1000).
'   - 28/nov/2008 lucianol acerto no gerador para criação de estrutura de grants para usuários.
'   - 11/jan/2008 lucianol criação componente flash com possibilidade de popup.
'   - 11/jan/2008 lucianol alteração de função CarregaEvento para possibilitar programação tanto no SAFARI quanto no Microsoft, incluindo attachEvent e obj[evento]=func.
'   - 11/jan/2008 lucianol inclusão função AdicionaFuncao(obj, evento, funcao) para permitir incluir no SAFARI uma função sem cancelar as anteriores(concatenar funções).
'   - 11/jan/2008 lucianol ScrollLeft e ScrollTop para resolver problema de obtenção de posição de scroll, que é diferente entre os navegadores.
'   - 11/jan/2008 lucianol Centraliza com código suficiente para posicionar elemento no centro da tela.
'   - 11/jan/2008 lucianol criado componente LightBox para facilitar inclusão do recurso em javascript.
'   - 11/jan/2008 lucianol alteração da rotina de IncluiStyleSheet para pegar no diretório inc por default.
'   - 11/jan/2008 lucianol alteração do LightBox, inclusão de opção de grupo para permitir anterior e posterior.
'   - 13/jan/2008 lucianol inclusão do prototype.js entre as bibliotecas carregadas pela intercraft por ser exigido pelo lightbox.
'   - 13/jan/2008 lucianol alteração da função $() no javascript icraft.js para $_() por já existir no prototype e alteração de todas para uso desta nova.
'   - 13/jan/2008 lucianol alteração $_() para considerar param "window" e "document", retornando controle obj correspondente.
'   - 14/jan/2008 lucianol alteração $_() para ajustar no do campo trocando "$" para "_", pois uniqueID retorna nome com "$".
'   - 14/jan/2008 lucianol função DebugPrint alterada para ao invés de dar erro, apresentar [[erro]] quando ocorre a tentativa de impressão de conteúdo de propriedade inexistente.
'   - 19/jan/2008 lucianol alteração na classe html incluindo PROTEGE para retornar apenas html previsível sem scripts etc
'   - 22/jan/2008 lucianol rotina paginacao permitindo posicionamento não só pelo num de página como também por chave, sendo esta passada em arraylist (paramtoarraylist)
'   - 22/jan/2008 lucianol inclusão da rotina DSDataColumns("campo1;campo2") para facilitar definição de primarykey em dataset
'   - 11/abr/2009 lucianol NZ, tornando ero também condição de retorno default. Muito trabalhoso e constante os tratamentos de erro com rotinas que envolvem propriedades opcionais. Nestes casos, NZ também poderá ser utilizado.
'   - 11/abr/2009 lucianol Prop, função não previa preenchimento de ImageUrl. Acertado.
'   - 11/abr/2009 lucianol ExibeData, tive problemas com data retornada nula. Para evitar situações como esta, ativei uma série de formatos:          PADRÃO "dd/MM/YYYY HH:mm:ss", "dd de mmmm de yyyy", "i", "a", "ai", "mmm dd, yyyy", "mmm dd, yyyy i", "mmm dd, yyyy c", "mmmm, yyyy i", "mmmm, yyyy c", "dd de mmmm de yyyy c" entre outros.
'   - 18/abr/2009 lucianol email/valida início de estrutura para operar objetos no sistema
'   - 19/abr/2009 lucianol enviaemail - inclui rotina de incorporação de imagens que faz uso de diretório temporário.
'   - 19/abr/2009 lucianol enviaemail - inclui rotina de autenticação em smtp
'   - 19/abr/2009 lucianol enviaemail - inclui rotina para evitar envio de estruturas já tratadas enviaemail(email,smtp...)
'   - 20/abr/2009 lucianol abstrcarac - para permitir conferência tipo dígito verificador
'   - 20/abr/2009 lucianol base36 - que retorna base36 de um número 0-9 A-Z limitando em nr casas considerando menos significativas
'   - 20/abr/2009 lucianol base36alga - transforma num em um dígito sendo a partir de 35 "z"
'   - 21/abr/2009 lucianol prop - inclusão de retorno name ao invés de id quando objeto é um control de windows forms
'   - 21/abr/2009 lucianol toda biblioteca - mudança completa de todos os controlcollections e controls para object permitindo utilização tanto em app quanto em web.ui
'   - 21/abr/2009 lucianol strstr - correção em critério de retorno com base em primeiro parâmetro negativo
'   - 26/abr/2009 lucianol enviaemail - alteração da sobrecarga de controle total (params retornados por byref) para retornar arquivos temporários caso sejam necessários, cids entre outros params. Exemplo:
'
'                           Dim mail As MailMessage = Nothing ' obriga que primeira execução inicie a mensagem
'                           Dim smtp As SmtpClient = Nothing ' obriga que primeira execução carregue smtp correto
'                           Dim tmps As New ArrayList ' retornará arquivos temporários caso sejam utilizados
'                           Dim cids As New ArrayList ' retornará cids caso imagens incorporadas
'                           Dim ret As New System.Text.StringBuilder
'                           ret.AppendLine(Icraft.EnviaEmail(mail, smtp, "lucianol@icraft.com.br", "lucianol@icraft.com.br", "teste com figura attach", "<img src=""http://www.intercraft.inf.br/figuras/bd02.jpg""/>", MailPriority.High, "smtpi.icraft.com.br", , , , , , True, cids, tmps, ParamArrayToArrayList("http://www.intercraft.inf.brx/", "\\webserver\inetpub\Intercraft\")))
'                           ret.AppendLine(Icraft.EnviaEmail(mail, smtp, , "luciano.lisboa@intermesa.com.br"))
'                           MsgBox(ret.ToString) ' mostra retorno
'                           mail.Dispose() ' limpa email para liberar os arquivos
'                           Icraft.ApagaTemps(tmps) ' apaga arquivos temporários
'
'   - 26/abr/2009 lucianol apagatemps - função que apaga arquivos mencionados no arraylist tmps
'   - 26/abr/2009 lucianol listadir - retorna arraylist contendo lista de arquivos disponíveis no diretório especificado
'   - 07/mai/2009 lucianol criação de dll a partir do icraft.vb. troca de nome da classe principal para icftbase e inclusão desta no namespace icraft.
'   - 18/jun/2009 lucianol ctypestr - tratando "on" como critério para booleano
'   - 18/jun/2009 lucianol prop - busca de conteúdo do campo booleano diretamente do formulário
'   - 28/jun/2009 lucianol infra - substituição icftmessage antigo pelo recurso feito em ajax
'   - 15/jul/2009 lucianol imagempath - função padronizada para busca de imagens em ambiente web
'   - 15/jul/2009 lucianol themepath - função para localizar arquivo em diretório de tema
'   - 15/jul/2009 lucianol ajuste no formato de header do arquivo padronizado para toda solução icraft
'   - 18/jul/2009 lucianol exibehtml - alteração da função para permitir definição de [link:url|descrição], [imgbut:urlimg|urllink|descrição] e [img:url|legenda]
'   - 18/jul/2009 lucianol exibehtml - classe para incorporar recursos exibehtml, temporária ainda, mas com as funções de handle para replace por regex
'   - 18/jul/2009 lucianol htmlreplimgbut - para regex replace ref ao exibe [imgbut:...]
'   - 18/jul/2009 lucianol htmlreplink - para regex replace ref ao exibe [link:...]
'   - 18/jul/2009 lucianol htmlreplimg - para regex replace ref ao exibe [img:...]
'   - 18/jun/2009 lucianol exibehtmlenc - para tratamento e encapsulamento de html
'   - 19/jul/2009 lucianol incluicampo - formatação de campo tipo data com tamanho menor por causa do assistente de calendário
'   - 19/jul/2009 lucianol incluicampo - especificação de tamanho para compobox, pois não estava sendo definido
'   - 19/jul/2009 lucianol imageurl - para obter diretório da imagem. caso não seja especificado, será avaliado como img
'   - 19/jul/2009 lucianol imagearq - para retornar diretório em disco de imagem específica
'   - 19/jul/2009 lucianol regexpamostra - bloqueio dos caracteres '<' e '>' para evitar tratamento de html
'   - 16/ago/2009 lucianol textologex - em web, incluir também variáveis de sessão no texto correspondente ao erro
'   - 16/ago/2009 lucianol logonsession - tostring para apresentar dados do logon de usuário
'   - 16/ago/2009 lucianol nz - preparo para conversão de logonsession para string, apresentando suas informações no formato (atrib=...;atrib2=...)
'   - 16/ago/2009 lucianol gerador microsoftxoracle - alteração do tipo float que deu problemas no fill do vb - single para number(8,6) e double para number(16,12)
'   - 16/ago/2009 lucianol gerador oraclexmicrosoft - tam > 16 ou decim > 2 para double
'   - 17/ago/2009 lucianol emailstr - obrigatoriedade de email no formato mínimo xxx@xxx.x
'   - 19/ago/2009 lucianol estadosdobrasil - retorna array com ufs do brasil
'   - 19/ago/2009 lucianol listadepaises - retorna array com nome de países por todo o mundo
'   - 19/ago/2009 lucianol exprexpr - rotina para concatenar expressões sem repetição de delimitador
'   - 19/ago/2009 lucianol fileexpr e urlexpr - utilização da função exprexpr
'   - 20/ago/2009 lucianol textologex - tanto para http como apl ignora mensagem de erro caso seja nothing
'   - 24/ago/2009 lucianol tiraacento - rotina que elimina acento do texto
'   - 24/ago/2009 lucianol obtempag - obtém página da internet em texto
'   - 24/ago/2009 lucianol lprop - configurar retorno de texto para tentar obter conteúdo de objeto antes de buscar request
'   - 28/ago/2009 weslleya capitalizar - criação da função que capitalizar uma string passada como parâmetro
'   - 31/ago/2009 weslleya primletramaius - faz um tipo de capitalização especial adequado à função lprimletramaius desenvolvida em Oracle
'   - 03/set/2009 lucianol soemailstr - rotina que obtém trecho de email existente ou string vazia caso não exista email válido
'   - 16/set/2009 lucianol class email - incluindo rotina de busca de somente endereço e descrição
'   - 16/set/2009 lucianol paginacao - inclusão de suporte para datarowcollection (sem possibilidade de add)
'   - 16/set/2009 weslleya dsgrava - preparação de variáveis para gravação de detalhes de conexão de internet no Oracle
'   - 16/set/2009 weslleya gravaoraclerestr - alteração da geração de script do trigger para consumir as variáveis de detalhes de conexão de internet
'   - 26/set/2009 lucianol lprop - considerando on e off para gravação de valor booleano
'   - 31/out/2009 lucianol regexmasctags - para retornar máscara capaz de quebrar tags html de forma recursiva
'   - 31/out/2009 lucianol regexhtml - manipulador de texto como html prox, dentro etc
'   - 31/out/2009 lucianol entifica - troca caracteres especiais de texto por códigos de entidade html
'   - 02/nov/2009 lucianol debugprint - para facilitar diagnóstico em telas de servidor
'   - 09/nov/2009 danielcosta - DSCarrega para banco SQLServer
'   - 09/nov/2009 danielcosta - DSFiltra - Filtra o conteudo de um dataset auxilia na redução do número de acessos ao banco
'   - 27/nov/2009 lucianol email - inclusão da propriedade domínio
'   - 28/nov/2009 lucianol chamaasync - inclusão de rotina delegada para chamada assíncrona, utilizando thread e liberando processador
'   - 02/dez/2009 weslleya CriadorDeObjetos - Criação da classe responsável por carregar dll e criar objetos diretamente dela
'   - 02/dez/2009 weslleya DsCarrega - Modificação dos objetos MySql para utilizar a classe CriadorDeObjetos
'   - 02/dez/2009 weslleya DsCarregaEstr - Modificação dos objetos MySql para utilizar a classe CriadorDeObjetos
'   - 02/dez/2009 weslleya DsGrava - Modificação dos objetos MySql para utilizar a classe CriadorDeObjetos
'   - 02/dez/2009 weslleya DsCriaComandoMySql - Modificação dos objetos MySql para utilizar a classe CriadorDeObjetos
'   - 02/dez/2009 weslleya DsCarrega - Modificação dos objetos Oracle para utilizar a classe CriadorDeObjetos
'   - 02/dez/2009 weslleya DsCarregaEstr - Modificação dos objetos Oracle para utilizar a classe CriadorDeObjetos
'   - 02/dez/2009 weslleya DsGrava - Modificação dos objetos Oracle para utilizar a classe CriadorDeObjetos
'   - 02/dez/2009 weslleya DsCriaComandoOracle - Modificação dos objetos Oracle para utilizar a classe CriadorDeObjetos
'   - 25/dez/2009 lucianol exibedata - ao invés de month, acertei para utilizar month - 1, pois array começa do zero
'   - 25/dez/2009 lucianol listadir - inclusão de critério para pesquisa de arquivos
'   - 26/dez/2009 lucianol exibehtml - retorno de html quando iniciado e finalizado com html
'   - 28/dez/2009 lucianol vardesessao - forma padronizada de montar variável para inclusão em sessão
'   - 28/dez/2009 lucianol atrib - criação de função para simplificar processo de implementação de atributos em todos os componentes
'   - 28/dez/2009 lucianol dscarrega e diversos - possibilidade de passar string de conexão diretamente com params providerName:System.Data.OleDb;Provider:Microsoft.Jet.OLEDB.4.0;Data Source:~/UC/ICFTCOMBOBOX/TESTE/ICFTCOMBOTESTE.MDB, sendo ~ substituído pela raiz do diretório
'   - 28/dez/2009 lucianol combobox - correção de rotina de atualização, que considerava incorretamente attribute atualizar ao invés de prope atualizar
'   - 28/dez/2009 lucianol combobox - inclusão de conceito excluirvalores que relaciona valores que não devem ser incluídos no combo (caso mais de uma coluna, considerar primeira como valor - chave composta deverá ser concatenada evitando repetição)
'   - 29/dez/2009 weslleya exprexpr - correção da função para substituir DelimAlternativo por Delim
'   - 30/dez/2009 lucianol itemtoarraylist e itemtoobject - busca por atrib, depois por item depois por attribute
'   - 30/dez/2009 lucianol gerador msaccess - inclusão de descrições em tabelas de sistema
'   - 30/dez/2009 lucianol gerador oracle - proteção de código evitando nothing em visoes e usuários
'   - 30/dez/2009 carandre gravaoraclesemrestr - acerto de função for cur em user job passando de owner para schema_user
'   - 30/dez/2009 carandre gravaoraclesemrestr - inclusão de comentário na função que gera script tanto para todas as tabelas inclusive as do sistema
'   - 30/dez/2009 lucianol pegatitulopagwebdocabeca - para obter titulo que se encontra no cabeçalho da página
'   - 30/dez/2009 lucianol pegatitulopagweb - para pegar título entre as tags title no header
'   - 30/dez/2009 lucianol pegahtmlemarquivo - para pegar htmlregex em arquivo 
'   - 30/dez/2009 lucianol fileexpr - inclusão de tratamento de ~/ para referenciar à home do site
'   - 30/dez/2009 lucianol imageurl - correção de probl. estava com fileexpr ao invés de urlexpr
'   - 30/dez/2009 lucianol listadir - troca de verificação de ~/ pelo fileexpr
'   - 01/jan/2010 lucianol obtemtexto - rotina para obter texto a partir de diretório ou url
'   - 08/jan/2010 weslleya novasenha - incorporação da função novasenha criada por Luciano na biblioteca
'   - 08/jan/2010 anderson exibedata - implementação do formato mmmm/yy para datas do tipo Dezembro/09
'   - 09/jan/2010 lucianol classe gerador - inclusão de propriedade xml para permitir entrara e obtenção de estrutura
'   - 09/jan/2010 lucianol classe campo - inclusão da propriedade tabela na classe campo para permitir salvamento a partir da classe gerador
'   - 09/jan/2010 lucianol tipocomotabela - para transformar uma classe em tabela e propriedades de classe em campos
'   - 09/jan/2010 lucianol textoemstream - para retornar um stream com um texto específico fazendo uso de memorystream
'   - 09/jan/2010 lucianol class gerador - inclusão de parâmetro GeraTabsSistema para indicar que o usuário deseja gerar tabelas do sistema
'   - 09/jan/2010 lucianol class gerador - inclusão de variável tabssistema contendo lista de tabelas criadas automaticamente pelo sistema
'   - 09/jan/2010 lucianol carregamsaccess - acerto de descrições para tabelas do sistema, inclusão de critério geratabssistema
'   - 09/jan/2010 lucianol classe tnsnamesreader para obter nomes de serviços do oracle
'   - 17/jan/2010 lucianol soma - rotina para somar números de uma sequência de valores
'   - 17/jan/2010 lucianol class gerador - inclusão de critérios alterando apenas a rotina de carga a partir do oracle, por falta de tempo
'   - 20/jan/2010 lucianol classe form - inclusão de rotina buscatipo para retornar um determinado tipo procurado por todo sistema de objetos
'   - 22/jan/2010 lucianol substituição do critério gerasistema pelo exporta infrasistema
'   - 31/jan/2010 lucianol filtrocampoconteudo - para montar filtro de sql com base nos campos de estrutura de VO e conteúdos
'   - 02/fev/2010 lucianol atualização tabelas do gerador inclusão de usuário e direito
'   - 02/fev/2010 lucianol tipoaccessToscript - retirada da função de dentro da classe do gerador para colocá-la diretamente no icftbase
'   - 02/fev/2010 lucianol acessook - inclusão de função para validação de login
'   - 02/fev/2010 lucianol masteracessook - função para verificação de login na masterpage
'   - 02/fev/2010 lucianol buscatipo - alteração para considerar tanto lista de tipos como string de tipos, permitindo procurar "system.string"
'   - 02/fev/2010 lucianol buscaprimeirotipo - retorna primeiro objeto daquele tipo de forma direta (sem vetor)
'   - 02/fev/2010 lucianol buscatipo - inclusão de javerificado para não entrar em controle já analisado
'   - 02/fev/2010 lucianol registracontrolecomopostback - inclui controle na lista de postback do primeiro updatepanel existente
'   - 09/fev/2010 lucianol emailstr - considerar além do espaço caractere (128+32=160) como espaço
'   - 09/fev/2010 lucianol htmld - html decode especial. retira todas as tags html
'   - 09/fev/2010 lucianol tirahtml - retira tags html
'   - 10/fev/2010 lucianol itemencode - configura conteúdo de forma que não possua códigos como ponto e vírgula e dois pontos, utilizados como separadores em expressões Icraft
'   - 10/fev/2010 lucianol itemdecode - rotina inversa ao itemencode, que retorna conteúdo previamente codificado
'   - 10/fev/2010 luicanol salvacontroles - para armazenamento de conteúdo de controles
'   - 10/fev/2010 lucianol recuperacontroles - para recuperação de conteúdo de controles a partir do salvamento prévio
'   - 18/fev/2010 lucianol form.controles - com parâmetro para permitir busca de controles de forma hierarquica (recursiva) permitindo especificação de mais de um prefixo
'   - 18/fev/2010 lucianol recuperadoform - para obter dados salvos em page.request.form conforme prefixos
'   - 18/fev/2010 lucianol urlexpr - converte c:\... em ~/ caso esteja na raiz do site
'   - 20/fev/2010 lucianol prop - vai tentar recuperar via reflexion antes de recuperar attributes no caso de propriedade não identificada pelo código
'   - 20/fev/2010 lucianol nz - conversão do tipo enum com ctype retornava número. caso seja enum, utilizará formato enum.tostring para retornar texto específico
'   - 20/fev/2010 lucianol incluicampo - alteração de rotina para considerar linha em tabela ao invés de inclusão de campo em divs soltos
'   - 22/fev/2010 lucianol fileexpr - inclusão de substituição de raiz fazendo uso do resolveurl('~/')
'   - 22/fev/2010 lucianol obtercor - rotina para obter cor a partir de um texto específico
'   - 22/fev/2010 lucianol incluicampo - modificação do nome do calendário para considerar não só calenda como o nome do campo evitando erro de duplicidade no mesmo form
'   - 22/fev/2010 lucianol notamsg mostra e mostrasem - eliminadas estas funções por não serem mais necessárias (ativação do notamsg ajax)
'     22/fev/2010 lucianol prope - precisei acertar rotina, pois criando attributo com add, mas quando utilizava como tag texto, gerava falha na interpretação do ==
'     22/fev/2010 lucianol conteudo - salvamento considerando valores em itemencode para evitar problemas com caracteres especiais para tratamento de itemlista
'     22/fev/2010 lucianol aplicamascara - modificação de tooltip para organização de texto relativo ao caminho, salvasemcaminho e máscara
'     22/fev/2010 lucianol paginacao - consideração de dataset vazio se não possuir registro ou for nothing, pois não existia esta última opção
'     27/fev/2010 lucianol regaplkey - grava ou obtém atributo no regedit do software da máquina 
'     27/fev/2010 lucianol regmachinekey - grava ou obtém atributo no regedit do machine da máquina
'     28/fev/2010 lucianol dirreplica - classe para efetuar replica de arquivos em diretório e subdiretórios
'     28/fev/2010 lucianol segsexpr - função para formatar segundos em hora, minuto e segundos
'     28/fev/2010 lucianol exibesegs - exibe segundos em diversos formatos
'     28/fev/2010 lucianol pl - função para retorno de singular ou plural conforme um número especificado



'
' IDÉIAS/NECESSIDADES:::
'   - xtecnico xxcomentárioxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'   - lucianol necessária criação de sys_config_global também no access
'   - lucianol necessária atualização de conteudo para config em referências para campo de sys_config_global
'   - lucianol apagar arquivos desnecessários em solução lightbox como imagens, por exemplo.
'   - lucianol ferramenta capaz de verificar 
'   - lucianol arrayv, ver velocidade de busca no array por delegate e busca por for each.
'   - lucianol elementosstr, registrar velocidade e buscar instruções de par ordenado.
'   - lucianol carregacombo, implementar regex consideraando "teste de ; texto";"mais um".
'   - lucianol macrosubst, deixar claro os itens que sofrem interpretação de macrosubst: sql no carrega, tabela e sistema no combobox.
'   - lucianol macrosubst, verificar performance de consultas sem e consultas com macrosubst.
'   - lucianol nz, ver comparação de performance de if not isnothing(..) e nz.
'   - lucianol notamsg, pedir para pessoal fazer tela html para montagem em java (ou control.add) da mensagem de erro para usuário.
'   - lucianol notamsg, permitir parametrização de erros desconhecidos para facilitar testes (mediante a verificação, algumas redes poderiam efetuar o cadastro do regex correspondente e mensagem padronizada).
'   - lucianol notamsg, seria interessante criar um componente ao invés de rotina de biblioteca que tivesse layout e o verifica já no PRE_RENDER.
'   - lucianol itenstoarraylist, necessária utilização da função prop(variavel, propriedade) para expandir a funcionalidade.
'   - lucianol notamsg, hoje, apresenta apenas msg e dá opção de fechar.
'   - lucianol notamsg, caso urldestino esteja vazia, será realmente a opção de fechar, mas caso esteja preenchida, ao invés de fechar, apresentará continuar sendo o javascript um redirecionamento para esta nova página..
'   - lucianol notamsg, bloquear os eventos do javascript enquanto a msg estiver sendo apresentada (simular popup)..
'   - lucianol showjsmessage, incluir parâmetro na função javascript showjsmessage permitindo que se passe um texto para título da janela..
'   - lucianol notamsg, incluir enum que apresente os diversos estados (ou títulos) facilitando a string "carregando registro", "carregando página"....
'   - lucianol proxseq, utilizado em chaves compostas (inclusão de filtro).
'   - lucianol valordefault, deve tratar macrosubstituição.
'   - lucianol now em format no macrosubstituição.
'   - lucianol ip e outros, request e server no macrosubst.
'   - lucianol logon, deve considerar ip (função de logon simples e concatenado com ip).
'   - lucianol notamsg, deve registrar erro como ocorre no errohttpmodule.
'   - lucianol form, muito difícil, mas seria interessante marcar conteúdo pesquisado em campos (amarelo).
'   - lucianol detailsview, atualizar somente mediante alguma atualização.
'   - lucianol gridview e detailsview altera atualizar para mudancadeselecao
'   - lucianol     e mudancadedados :selecao para alteracao de chave, o que
'   - lucianol     obrigaria reposicionamentos das dependências / :mudancadedados 
'   - lucianol     para atualizacao dos dados em dependências.
'   - lucianol icftform, ao alterar combo tarefa usuário, campo cliente no filtro, ocorre postback sem necessidade.
'   - lucianol prop, campo booleano, ao buscar do formulário está alterando valor mesmo com variável inalterada

'
' --------------------------------------------------------------------------------
'
'
'
Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Mail
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Windows.Forms.Form
Imports System.Xml
Imports Microsoft.Win32

Imports System.Data.SqlClient

Namespace Icraft

    Public Class IcftBase

        ''' <summary>
        ''' Separador padrão de colunas em combo ou elementos da biblioteca de um modo geral: PARAM1 | PARAM2 ...
        ''' </summary>
        ''' <remarks></remarks>
        Public Const ComboSepDefault As String = " | "

        ''' <summary>
        ''' String que aparece na combo quando item vazio é incluído, representando nenhuma seleção.
        ''' </summary>
        ''' <remarks></remarks>
        Public Const ComboNull As String = "--"

        ''' <summary>
        ''' Definições DAO para permitir vínculo conforme necessidade.
        ''' </summary>
        ''' <remarks></remarks>
        Public Const DAO_RelationAttributeEnum_dbRelationDeleteCascade As Integer = 4096
        Public Const DAO_RelationAttributeEnum_dbRelationUpdateCascade As Integer = 256
        Public Const DAO_RelationAttributeEnum_dbRelationDontEnforce As Integer = 2
        Public Const DAO_DataTypeEnum_dbText As Integer = 10
        Public Const DAO_DataTypeEnum_dbDouble As Integer = 7
        Public Const DAO_DataTypeEnum_dbDate As Integer = 8
        Public Const DAO_DataTypeEnum_dbMemo As Integer = 12
        Public Const DAO_DataTypeEnum_dbBinary As Integer = 9
        Public Const DAO_DataTypeEnum_dbBoolean As Integer = 1
        Public Const DAO_LanguageConstants_dbLangGeneral As String = ";LANGID=0x0409;CP=1252;COUNTRY=0"

        ''' <summary>
        ''' Tipos de bases previstas, nomes de providers para facilitar.
        ''' </summary>
        ''' <remarks></remarks>
        Public Const MySQL As String = "MySql.Data.MySqlClient"
        Public Const MSAccess As String = "System.Data.OleDb"
        Public Const Oracle As String = "System.Data.OracleClient"
        Public Const SQLServer As String = "System.Data.SqlClient"

        ''' <summary>
        ''' Opções de comando para funções de formatação de data para gravação em banco de dados.
        ''' </summary>
        ''' <remarks></remarks>
        Enum TipoBaseSQL
            Gerador
            MSAccess
            MySQL
            Oracle
            SQLServer
            XML
        End Enum

        ''' <summary>
        ''' Opções de comando para funções de tratamento de SQL.
        ''' </summary>
        ''' <remarks></remarks>
        Enum ExprSQLTipo
            Sel
            Ins
            Upd
            Del
        End Enum

        ''' <summary>
        ''' Chave default utilizada por encrypt/decrypt para codificação de texto.
        ''' </summary>
        ''' <remarks></remarks>
        Const EncrypBChavePadrao As String = "AoPu2.%X´A¨'AÇç~^.M<"

        ''' <summary>
        ''' Idiomas padronizados facilitando seu uso em biblioteca.
        ''' </summary>
        ''' <remarks></remarks>
        Enum TipoIdioma
            PT_BR
            EN
            ES
        End Enum

        ''' <summary>
        ''' Opções para interpretação do elemento quando adicionado ao conjunto elementosstr.
        ''' </summary>
        ''' <remarks></remarks>
        Enum ElementoStrOpera
            Atribui
            Aumenta
            Diminui
        End Enum

        ''' <summary>
        ''' Texto encontrado em tipo de componente.
        ''' </summary>
        ''' <remarks></remarks>
        Const TipoTxtIcftMenu As String = "ASP.uc_icftmenu_ascx"

        ''' <summary>
        ''' Texto utilizado como tabulador em visualizações das comparações de estrutura no GERADOR.
        ''' </summary>
        ''' <remarks></remarks>
        Const Gerador_Tabula As String = "+--"

        ''' <summary>
        ''' Atributos previstos para procura node.
        ''' </summary>
        ''' <remarks></remarks>
        Enum NodeCampo
            NavigateUrl
            Text
            TooolTip
            ValuePath
        End Enum

        ''' <summary>
        ''' Tipo de opção para execução NOCACHE.
        ''' </summary>
        ''' <remarks></remarks>
        Enum NoCacheTipo
            PaginaExpirada
            SemHistorico
        End Enum

        ''' <summary>
        ''' Automatizar combo significa garantir que as dependências (param ATUALIZAR) sejam atualizadas mediante qualquer alteração.
        ''' </summary>
        ''' <param name="ComboOuContainer">Deve ser um combobox ou um container, sendo todos combos existentes neste automatizados.</param>
        ''' <param name="CarregarDeUmaVez">Pode-se passar NOT POSTBACK garantir a carga somente ao ser carregada a página.</param>
        ''' <remarks></remarks>
        Shared Sub AutomatizaCombo(ByVal ComboOuContainer As Object, Optional ByVal CarregarDeUmaVez As Boolean = False)
            If TypeOf ComboOuContainer Is DropDownList Then
                If PropE(ComboOuContainer, "Atualizar") <> "" Then
                    AddHandler CType(ComboOuContainer, DropDownList).SelectedIndexChanged, AddressOf AtualizouControle
                    CType(ComboOuContainer, DropDownList).AutoPostBack = True
                End If
                If CarregarDeUmaVez Then
                    CarregaCombo(ComboOuContainer)
                End If
            Else
                For Each Opc As Object In Form.Containers(ComboOuContainer)
                    For Each Ctl As Control In Opc.controls
                        If TypeOf Ctl Is DropDownList Then
                            If PropE(Ctl, "Atualizar") <> "" Then
                                AddHandler CType(Ctl, DropDownList).SelectedIndexChanged, AddressOf AtualizouControle
                                CType(Ctl, DropDownList).AutoPostBack = True
                            End If
                            If CarregarDeUmaVez Then
                                CarregaCombo(Ctl)
                            End If
                        End If
                    Next
                Next
            End If
        End Sub

        ''' <summary>
        ''' Carrega o combobox conforme parâmetros: SQL, QtdCols, OcultaValor, StrConn entre outros, definidos como atributos do controle passado.
        ''' </summary>
        ''' <param name="Combo">Combo a ser carregado.</param>
        ''' <remarks>
        ''' PARAMS
        '''        [Atualizar] = corresponde à campos separados por ponto e vírgula que serão atualizados automaticamente mediante alteração do valor deste combo.
        '''        [QtdCols] = quantidade de colunas
        '''        [OcultaValor] = true para não apresentar valor na expressão de apresentação
        '''          considerado somente quando mais de uma coluna.
        '''          importante saber que a primeira coluna sempre será a de valor, podendo ou não ser apresentada, conforme parâmetro ocultacoluna.
        '''        [SQL] = cláusula de select a ser submetida, podendo conter [:CAMPO] para pesquisa em todo o formulário, containers irmãos, filhos e depois pais até PAGE.
        '''        [StrConn] = string de conexão.
        '''        [SeparaCols] = separador de colunas.
        '''        [NotNull] = não inclui valor que representa o nulo.
        ''' </remarks>
        Shared Sub CarregaCombo(ByRef Combo As DropDownList)
            ' verifica se existe campos para atualizar
            Dim Lista As ArrayList = Nothing
            If PropE(Combo, "SQL") <> "" Then
                Lista = ParamArrayToArrayList(DSCarrega(PropE(Combo, "SQL"), PropE(Combo, "StrConn"), Combo.Parent, Logon(Combo.Page)))
            ElseIf PropE(Combo, "Lista") <> "" Then
                Lista = ParamArrayToArrayList(Split(PropE(Combo, "Lista"), ";"))
            End If

            Dim ExcluirValores As String = PropE(Combo, "EXCLUIRVALORES")
            Dim QtdCols As Integer = NZV(PropE(Combo, "QtdCols"), "1")

            If ExcluirValores <> "" Then
                Dim z As Integer = 0
                Do While z < Lista.Count
                    If TemNaLista(ExcluirValores, Lista(z).ToString) Then
                        For zz = 0 To QtdCols - 1
                            Lista.RemoveAt(z)
                        Next
                    Else
                        z += QtdCols
                    End If
                Loop
            End If

            If Not IsNothing(Lista) Then
                If Not CType(NZV(PropE(Combo, "NotNull"), Boolean.FalseString), Boolean) Then
                    CarregaCombo(Combo, True, ComboNull)
                Else
                    Combo.Items.Clear()
                End If

                If QtdCols = 1 Then
                    CarregaCombo(Combo, False, Lista)
                Else
                    CarregaCombo(Combo, CType(QtdCols, Integer), CType(NZV(PropE(Combo, "OcultaValor"), Boolean.FalseString), Boolean), NZV(PropE(Combo, "SeparaCols"), ComboSepDefault), False, Lista)
                End If
            End If
        End Sub

        ''' <summary>
        ''' Carrega valores em um combobox. Cada valor da lista será uma linha do combo.
        ''' </summary>
        ''' <param name="Combo">Combobox a ser carregado.</param>
        ''' <param name="Limpa">True limpa (clear) antes ou false para acrescentar itens.</param>
        ''' <param name="Params">Lista de parâmetros, que pode ser string, array, arraylist ou dataset.</param>
        ''' <remarks></remarks>
        Shared Sub CarregaCombo(ByVal Combo As Object, ByVal Limpa As Boolean, ByVal ParamArray Params() As Object)
            CarregaCombo(Combo, 1, False, "", Limpa, Params)
        End Sub


        ''' <summary>
        ''' Carrega combo com expressões (textos apresentados) diferentes dos valores.
        ''' Os parâmetros são passados em sequência, sendo considerados além das strings, arraylists, vetores e datasets.
        ''' Valor sempre será a primeira coluna.
        ''' </summary>
        ''' <param name="Combo">Combo a ser atualizado.</param>
        ''' <param name="QtdCols">Quantidade de colunas existentes nos parâmetros sequenciados.</param>
        ''' <param name="OcultaValor">Vai apresentar como início da expressão a coluna zero ou não.</param>
        ''' <param name="SeparadorCols">Separador de colunas, caso diferente do padrão.</param>
        ''' <param name="Limpa">True se o combo será limpo antes ou false para acrescentar valores.</param>
        ''' <param name="Params">Lista de parâmetros, que podem ser strings, arrays, arraylists ou datasets.</param>
        ''' <remarks>
        ''' Exemplo1:
        '''     select matr, nome from sócio
        '''     são duas colunas (QtdCols=2)
        '''     mostrando também a matrícula (OcultaValor=false)
        '''     sem a preocupação com separador de colunas
        ''' 
        ''' Exemplo2:
        '''     select cod, CPF, bairro, cidade, estado from clente
        '''     são 5 colunas (QtdCols=5)
        '''     não mostraremos o código (OCultaValor=true)
        '''     colocando separador como ", " (SeparadorCols=", ")
        ''' </remarks>
        Shared Sub CarregaCombo(ByRef Combo As Object, ByVal QtdCols As Integer, ByVal OcultaValor As Boolean, ByVal SeparadorCols As String, ByVal Limpa As Boolean, ByVal ParamArray Params() As Object)

            ' transforma todos os itens em parâmetros serializáveis
            Dim Itens As ArrayList = ParamArrayToArrayList(Params)

            ' caso solicitado, limpa o combo
            If Limpa Then
                Combo.Items.Clear()
            End If

            ' separador
            If SeparadorCols = "" Then
                SeparadorCols = ComboSepDefault
            End If

            ' primeira coluna sempre com o valor.
            ' caso já exista no combo, não coloca denovo...
            For z As Integer = 0 To Itens.Count - 1 Step QtdCols
                Dim ListIT As ListItem
                If QtdCols > 1 Then
                    Dim Expr As String = ""
                    For zz As Integer = (z + IIf(OcultaValor, 1, 0)) To z + QtdCols - 1
                        Expr &= IIf(Expr <> "", SeparadorCols, "") & NZ(Itens(zz), "")
                    Next
                    ListIT = New ListItem(Expr, Itens(z))
                Else
                    ListIT = New ListItem(NZ(Itens(z), ComboNull))
                End If

                If Not Combo.Items.Contains(ListIT) Then
                    Combo.Items.Add(ListIT)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Rotina de atualização de combos fazendo uso de parâmetros como SQL, QTDCOLS etc diretamente.
        ''' </summary>
        ''' <param name="Combo">Combobox a ser atualizado.</param>
        ''' <param name="SQL">Expressão SQL podendo conter termos de marcrosubst como [:CAMPOFORM].</param>
        ''' <param name="StrConn">String de conexão utilizada para pesquisa.</param>
        ''' <param name="Atualizar">Campos a serem atualizados por estes (sofrerão databind automático).</param>
        ''' <param name="VinculadoA">Campos que deverão atualizar este.</param>
        ''' <remarks></remarks>
        Shared Sub CarregaComboVinc(ByVal Combo As DropDownList, ByVal SQL As String, ByVal StrConn As String, Optional ByVal Atualizar As String = "", Optional ByVal VinculadoA As String = "")
            CarregaComboVinc(Combo, 1, False, "", SQL, StrConn, Atualizar, VinculadoA)
        End Sub

        ''' <summary>
        ''' Carrega combobox vinculando-o aos outros. Uma maneira de fazer atribuição de parâmetros por código de maneira simplificada.
        ''' Já automatiza combos vinculados e este, se também possuir ATUALIZA.
        ''' </summary>
        ''' <param name="Combo">Combobox a ser carregado.</param>
        ''' <param name="QtdCols">Quantidade de colunas.</param>
        ''' <param name="OcultaValor">True para ocultar a primeira coluna da expressão, caso QtdCols > 1.</param>
        ''' <param name="SeparadorCols">Separador de colunas caso diferente do padrão.</param>
        ''' <param name="SQL">Texto SQL a ser interpretado podendo conter macrosubst tipo [:CAMPOFORM].</param>
        ''' <param name="StrConn">String de conexão para pesquisa.</param>
        ''' <param name="Atualizar">Campos que serão atualizados por este, caso existam.</param>
        ''' <param name="VinculadoA">Campos que deverão atualizar este.</param>
        ''' <remarks></remarks>
        Shared Sub CarregaComboVinc(ByVal Combo As DropDownList, ByVal QtdCols As Integer, ByVal OcultaValor As Boolean, ByVal SeparadorCols As String, ByVal SQL As String, ByVal StrConn As String, Optional ByVal Atualizar As String = "", Optional ByVal VinculadoA As String = "")

            ' separador
            If SeparadorCols = "" Then
                SeparadorCols = ComboSepDefault
            End If

            PropE(Combo, "SQL") = SQL
            PropE(Combo, "QtdCols") = QtdCols
            PropE(Combo, "OcultaValor") = OcultaValor
            PropE(Combo, "SeparaCols") = SeparadorCols
            PropE(Combo, "StrConn") = StrConn
            PropE(Combo, "Atualizar") = Atualizar

            If VinculadoA <> "" Then
                For Each CampoStr As String In Split(VinculadoA, ";")
                    Dim Ctl As Control = Form.FindControl(Combo.Page, CampoStr)
                    If Not IsNothing(Ctl) AndAlso TypeOf Ctl Is DropDownList Then
                        Dim Vinc As ArrayList = ParamArrayToArrayList(Split(PropE(Ctl, "Atualizar"), ";"))
                        If Not Vinc.Contains(Combo.ID) Then
                            Vinc.Add(Combo.ID)
                            PropE(Ctl, "Atualizar") = Join(Vinc.ToArray, ";")
                            AutomatizaCombo(Ctl)
                        End If
                    End If
                Next
            End If

            If Atualizar <> "" Then
                Dim Vinc As ArrayList = ParamArrayToArrayList(Split(PropE(Combo, "Atualizar"), ";"))
                For Each AtualizarNome As String In Split(Atualizar, ";")
                    If Not Vinc.Contains(AtualizarNome) Then
                        Vinc.Add(Combo.ID)
                    End If
                Next
                PropE(Combo, "Atualizar") = Join(Vinc.ToArray, ";")
                AutomatizaCombo(Combo)
            End If
            CarregaCombo(Combo)
        End Sub


        ''' <summary>
        ''' Lista de países para preenchimento de fontes de dados de controles combo e afins.
        ''' </summary>
        ''' <value></value>
        ''' <returns>Um array de strings onde cada elemento representa um país da lista.</returns>
        ''' <remarks></remarks>
        Shared ReadOnly Property ListaDePaises() As Array
            Get
                Return Split("Brasil,Afeganistão,África do Sul,Albânia,Alemanha,Algéria,Andorra,Angola,Anguilla,Antártida,Antígua e Barbuda,Antilhas Holandesas,Arábia Saudita,Argentina,Armênia,Aruba,Austrália,Áustria,Azerbaijão,Bahamas,Bahrain,Bangladesh,Barbados,Belarus,Bélgica,Belize,Benin,Bermuda,Bolívia,Butão,Bósnia-Herzegovina,Botsuana,Brunei,Bulgária,Burkina Faso,Burundi,Cabo Verde,Camboja,Camarões,Canadá,Casaquistão,Chade,Chile,China,Chipre,Colômbia,Comoros,Congo,Coréia do Norte,Coréia do Sul,Costa do Marfim,Costa Rica,Croácia,Cuba,Dinamarca,Djibouti,Dominica,El Salvador,Equador,Egito,Emirados Árabes Unidos,Eritréia,Espanha,Eslováquia,Eslovênia,Estados Unidos da América,Estônia,Etiópia,Fiji,Filipinas,Finlândia,França,Gabão,Gâmbia,Gana,Geórgia,Gibraltar,Granada,Grécia,Groelândia,Guadalupe,Guam,Guatemala,Guiana Francesa,Guiné,Guiné-Bissau,Guiné Equatorial,Guiana,Haiti,Honduras,Hong Kong,Hungria,Iêmen,Ilhas Cayman,Ilha Bouvet,Ilhas Cocos,Ilhas Cook,Ilhas costeiras dos EUA,Ilhas Costeiras dos EUA,Ilhas Faroe,Ilhas Heard e McDonald,Ilhas Mariana do Norte,Ilhas Marshall,Ilhas Natal,Ilha Norfolk,Ilha Pitcairn,Ilhas S. Georgia e S. Sandwich,Ilhas Salomão,Ilhas Svalbard e Jan Mayen,Ilhas Turks e Caicos,Ilhas Virgens,Ilhas Virgens Britânicas,Ilhas Wallis e Futuna,Índia,Indonésia,Islândia,Irã,Iraque,Irlanda,Israel,Itália,Iugoslávia (ex-),Jamaica,Japão,Jordânia,Kiribati,Kuwait,Kyrgyztan,Laos,Látvia,Lesoto,Líbano,Libéria,Líbia,Liechtenstein,Lituânia,Luxemburgo,Macau,Macedônia,antiga Iugoslávia,Madagascar,Malásia,Malaui,Maldivas,Mali,Malta,Marrocos,Martinica,Maurício,Mauritânia,Mayotte,México,Micronésia,Moçambique,Moldova,Mônaco,Mongólia,Montserrat,Myanmar,Namíbia,Nauru,Nepal,Holanda,Nicarágua,Niger,Nigéria,Niue,Noruega,Nova Caledônia,Nova Zelândia,Oman,Palau,Panamá,Papua Nova Guiné,Paquistão,Paraguai,Peru,Polinésia Francesa,Polônia,Porto Rico,Portugal,Qatar,Quênia,Reino Unido,República Centro-Africana,República Dominicana,República Tcheca,Reunião,Romênia,Ruanda,Rússia,Saara Ocidental,Saint Kitts e Nevis,Saint Vincent e Grenadines,Samoa,Samoa Americana,San Marino,Santa Helena,Santa Lúcia,São Tomé e Príncipe,Senegal,Serra Leão,Seychelles,Singapura,Síria,Somália,Sri Lanka,St. Pierre e Miquelon,Sudão,Suriname,Suazilândia,Suécia,Suíça,Tailândia,Taiwan,Tajikistão,Tanzânia,Territórios Franceses do Sul,Território marítimo das Índias Britânicas,Timor Leste,Togo,Tokelau,Tonga,Trinidad e Tobago,Tunísia,Turcomenistão,Turquia,Tuvalu,Ucrânia,Uganda,Uruguai,Usbequistão,Vanuatu,Vaticano,Venezuela,Vietnã,Zaire,Zâmbia,Zimbábue", ",")
            End Get
        End Property

        ''' <summary>
        ''' Lista de estados do Brasil para preenchimento de fontes de dados de controles combo e afins.
        ''' </summary>
        ''' <value></value>
        ''' <returns>Um array de strings onde cada elemento representa um estado do Brasil presente na lista.</returns>
        ''' <remarks></remarks>
        Shared ReadOnly Property EstadosDoBrasil() As Array
            Get
                Return Split("AC,AL,AM,AP,BA,CE,DF,ES,GO,MA,MG,MS,MT,PA,PB,PE,PI,PR,RJ,RN,RO,RR,RS,SC,SE,SP,TO,NA", ",")
            End Get
        End Property




        ''' <summary>
        ''' Envia email e, caso ocorra, retorna string de erro. Formato completo exige dispose do EMAIL e apagatemps no final.
        ''' </summary>
        ''' <param name="Mail">Caso já esteja pronto, poderá passar obj de mensagem ou retorná-lo após envio.</param>
        ''' <param name="De">Nome e email do remetente no formato: 'nome' [email@dominio.com.br].</param>
        ''' <param name="Para">Nome e email dos destinatários no formato: 'nome1' [email1@dominio.com.br];'nome2' [email2@dominio.com.br].</param>
        ''' <param name="Assunto">Texto de assunto da mensagem.</param>
        ''' <param name="Corpo">Corpo da mensagem em html.</param>
        ''' <param name="Prioridade">Nível de prioridade entre alta, normal e baixa.</param>
        ''' <param name="SmtpHost">Servidor de smtp. Na sua ausência, smtp_host do webconfig será considerado.</param>
        ''' <param name="SmtpPort">Porta de smtp. Na sua ausência, smtp_port do webconfig será considerada.</param>
        ''' <param name="CC">Com cópia. Deve conter lista 'nome' [email@dominio.com.br];'nome2' [email2@dominio.com.br].</param>
        ''' <param name="BCC">Com cópia oculta. Também pode ser informado com BCC na frente de qualquer destinatário. Deve conter lista 'nome' [email@dominio.com.br];'nome2' [email2@dominio.com.br].</param>
        ''' <param name="SMTPUsuario">Usuário de autenticação no SMTP.</param>
        ''' <param name="SMTPSenha">Senha de autenticação no SMTP.</param>
        ''' <param name="IncorporaImagens">Ordena incorporação das imagens ao invés de seguirem links para elas.</param>
        ''' <param name="UrlsLocais">Parâmetros contendo url e dir correspondente para redirecionamento, podendo ser mais de um par.</param>''' 
        ''' <returns>Retorna um texto correspondente ao erro ou "" caso o envio tenha sido um sucesso. Retorna também variáveis atualizadas: Smtp, Mail, Corpo, CIDS e TMPS. Se optar por este formato, deverá preocupar-se em dar dispose no mail e apagar arquivos temporários.</returns>
        ''' <param name="Attachs">Lista de arquivos a seguirem attachados.</param>
        ''' <remarks></remarks>
        Public Shared Function EnviaEmail(ByRef Mail As MailMessage, Optional ByRef Enviar As System.Net.Mail.SmtpClient = Nothing, Optional ByVal De As String = Nothing, Optional ByVal Para As Object = Nothing, Optional ByVal Assunto As String = Nothing, Optional ByRef Corpo As String = Nothing, Optional ByVal Prioridade As System.Net.Mail.MailPriority = Nothing, Optional ByVal SmtpHost As String = Nothing, Optional ByVal SmtpPort As Integer = 25, Optional ByVal CC As Object = Nothing, Optional ByVal BCC As Object = Nothing, Optional ByVal SMTPUsuario As String = Nothing, Optional ByVal SMTPSenha As String = Nothing, Optional ByVal IncorporaImagens As Boolean = False, Optional ByRef CIDS As ArrayList = Nothing, Optional ByRef TMPS As ArrayList = Nothing, Optional ByVal UrlsLocais As ArrayList = Nothing, Optional ByVal Attachs As ArrayList = Nothing) As String
            Try

                ' cada param só é definido caso esteja mencionado
                If IsNothing(Mail) Then
                    Mail = New MailMessage
                End If

                If Not IsNothing(De) Then
                    Dim DeLista As ArrayList = TermosStrToLista(De)
                    Mail.From = New MailAddress(EmailStr(DeLista(0)))
                End If

                If Not IsNothing(Para) Or Not IsNothing(CC) Or Not IsNothing(BCC) Then
                    Mail.Bcc.Clear()
                    Mail.CC.Clear()
                    Mail.To.Clear()
                End If

                If Not IsNothing(Para) Then
                    Dim ParaLista As ArrayList = TermosStrToLista(Para)
                    For Each ParaItem As String In ParaLista
                        If ParaItem.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                            Dim M As New Email(ParaItem.Substring(4))
                            Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                        Else
                            Mail.To.Add(New MailAddress(EmailStr(ParaItem)))
                        End If
                    Next
                End If

                If Not IsNothing(CC) Then
                    Dim CCLista As ArrayList = TermosStrToLista(CC)
                    For Each ParaItem As String In CCLista
                        If ParaItem.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                            Dim M As New Email(ParaItem.Substring(4))
                            Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                        Else
                            Mail.CC.Add(New MailAddress(EmailStr(ParaItem)))
                        End If
                    Next
                End If

                If Not IsNothing(BCC) Then
                    Dim BCCLista As ArrayList = TermosStrToLista(BCC)
                    For Each ParaItem As String In BCCLista
                        If ParaItem.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                            Dim M As New Email(ParaItem.Substring(4))
                            Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                        Else
                            Dim M As New Email(ParaItem)
                            Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                        End If
                    Next
                End If

                If Not IsNothing(Prioridade) Then
                    Mail.Priority = Prioridade
                End If

                If Not IsNothing(Assunto) Then
                    Mail.Subject = Assunto
                End If

                If Not IsNothing(Corpo) Then
                    Mail.AlternateViews.Clear()

                    If Not IncorporaImagens Then
                        Mail.IsBodyHtml = True
                        Mail.SubjectEncoding = System.Text.Encoding.GetEncoding("UTF-8")
                        Mail.BodyEncoding = System.Text.Encoding.GetEncoding("UTF-8")
                        Mail.Body = Corpo
                    Else

                        ' inicia variávies de retorno caso não estejam definidas
                        If IsNothing(TMPS) Then
                            TMPS = New ArrayList
                        End If
                        If IsNothing(CIDS) Then
                            CIDS = New ArrayList
                        End If

                        ' visão alternativa
                        Dim alt As AlternateView = AlternateView.CreateAlternateViewFromString("", System.Text.Encoding.UTF8, "text/plain")
                        Mail.AlternateViews.Add(alt)

                        Dim arrImagens As New ArrayList
                        Dim listaImagens As String = "|"

                        For Each src As Match In Regex.Matches(Corpo, "url\(['|\""]+.*['|\""]\)|src=[""|'][^""']+[""|']", RegexOptions.IgnoreCase)
                            If InStr(1, listaImagens, "|" & src.Value & "|") = 0 Then
                                arrImagens.Add(src.Value)
                                listaImagens &= src.Value & "|"
                            End If
                        Next


                        For indx As Integer = 0 To arrImagens.Count - 1
                            Dim cid As String = "cid:EmbedRes_" & indx + 1
                            Corpo = Corpo.Replace(arrImagens(indx), "src=""" & cid & """")
                            Dim img As String = Regex.Replace(arrImagens(indx), "url\(['|\""]", "")
                            img = Regex.Replace(img, "src=['|\""]", "")
                            img = Regex.Replace(img, "['|\""]\)", "").Replace("""", "")


                            ' redirecionamentos
                            If Not IsNothing(UrlsLocais) Then
                                For Z = 0 To UrlsLocais.Count - 1 Step 2
                                    Dim urlcomp As String = UrlsLocais(Z)
                                    If img.StartsWith(urlcomp, StringComparison.OrdinalIgnoreCase) Then
                                        img = img.Replace(urlcomp, UrlsLocais(Z + 1))
                                    End If
                                Next
                            End If

                            Dim URL As New System.Uri(img)
                            If URL.Scheme = "http" Or URL.Scheme = "ftp" Then
                                ' carrega imagens caso remotas
                                Dim request As HttpWebRequest = WebRequest.Create(URL)
                                request.Timeout = 5000 ' cinco segundo de carga, senão erro...
                                Dim response As HttpWebResponse = request.GetResponse()
                                Dim bmp As New Bitmap(response.GetResponseStream)
                                img = System.IO.Path.GetTempFileName()
                                TMPS.Add(img)
                                bmp.Save(img)
                            End If
                            CIDS.Add(img)

                        Next

                        ' incorpora imagens
                        alt = AlternateView.CreateAlternateViewFromString(Corpo, System.Text.Encoding.UTF8, "text/html")
                        For z = 0 To CIDS.Count - 1
                            Dim res As New LinkedResource(CType(CIDS(z), String))
                            res.ContentId = "EmbedRes_" & z + 1
                            alt.LinkedResources.Add(res)
                        Next
                        Mail.AlternateViews.Add(alt)
                    End If
                End If

                ' inclui attachados
                If Not IsNothing(Attachs) Then
                    For Each attach As String In Attachs
                        Mail.Attachments.Add(New Attachment(attach))
                    Next
                End If

                If IsNothing(Enviar) Then
                    Enviar = New System.Net.Mail.SmtpClient(NZ(SmtpHost, WebConf("smtp_host")), NZV(NZ(SmtpPort, WebConf("smtp_port")), 25))
                End If

                If Not IsNothing(SMTPUsuario) Then
                    Enviar.Credentials = New System.Net.NetworkCredential(SMTPUsuario, NZ(SMTPSenha, ""))
                End If

                Enviar.Timeout = 100000


                Enviar.Send(Mail)

                Return ""
            Catch ex As Exception
                Return MessageEx(ex, "Erro ao tentar enviar email")
            End Try
        End Function

        ''' <summary>
        ''' Envia email e, caso ocorra, retorna string de erro.
        ''' </summary>
        ''' <param name="De">Nome e email do remetente no formato: 'nome' [email@dominio.com.br].</param>
        ''' <param name="Para">Nome e email dos destinatários no formato: 'nome1' [email1@dominio.com.br];'nome2' [email2@dominio.com.br].</param>
        ''' <param name="Assunto">Texto de assunto da mensagem.</param>
        ''' <param name="Corpo">Corpo da mensagem em html.</param>
        ''' <param name="Prioridade">Nível de prioridade entre alta, normal e baixa.</param>
        ''' <param name="SmtpHost">Servidor de smtp. Na sua ausência, smtp_host do webconfig será considerado.</param>
        ''' <param name="SmtpPort">Porta de smtp. Na sua ausência, smtp_host do webconfig será considerada.</param>
        ''' <param name="CC">Com cópia. Deve conter lista 'nome' [email@dominio.com.br];'nome2' [email2@dominio.com.br].</param>
        ''' <param name="BCC">Com cópia oculta. Também pode ser informado com BCC na frente de qualquer destinatário. Deve conter lista 'nome' [email@dominio.com.br];'nome2' [email2@dominio.com.br].</param>
        ''' <param name="SMTPUsuario">Usuário de autenticação no SMTP.</param>
        ''' <param name="SMTPSenha">Senha de autenticação no SMTP.</param>
        ''' <param name="IncorporaImagens">Ordena incorporação das imagens ao invés de seguirem links para elas.</param>
        ''' <param name="UrlsLocais">Parâmetros contendo url e dir correspondente para redirecionamento, podendo ser mais de um par.</param>''' 
        ''' <param name="Attachs">Lista de arquivos a seguirem attachados.</param>
        ''' <returns>Retorna um texto correspondente ao erro ou "" caso o envio tenha sido um sucesso.</returns>
        ''' <remarks></remarks>
        Public Shared Function EnviaEmail(ByVal De As String, ByVal Para As Object, ByVal Assunto As String, ByVal Corpo As String, Optional ByVal Prioridade As System.Net.Mail.MailPriority = MailPriority.Normal, Optional ByVal SmtpHost As String = Nothing, Optional ByVal SmtpPort As Integer = 25, Optional ByVal CC As Object = Nothing, Optional ByVal BCC As Object = Nothing, Optional ByVal SMTPUsuario As String = "", Optional ByVal SMTPSenha As String = "", Optional ByVal IncorporaImagens As Boolean = False, Optional ByVal UrlsLocais As ArrayList = Nothing, Optional ByVal Attachs As ArrayList = Nothing) As String
            Dim Mail As New MailMessage
            Dim Enviar As New System.Net.Mail.SmtpClient(NZ(SmtpHost, WebConf("smtp_host")), NZ(SmtpPort, WebConf("smtp_port")))
            Dim TMPS As New ArrayList
            Dim Ret As String = EnviaEmail(Mail, Enviar, De, Para, Assunto, Corpo, Prioridade, SmtpHost, SmtpPort, CC, BCC, SMTPUsuario, SMTPSenha, IncorporaImagens, , TMPS, UrlsLocais)
            Mail.Dispose() ' libera arquivos
            ApagaTemps(TMPS)
            Return Ret
        End Function

        ''' <summary>
        ''' Grava texto em arquivo no disco.
        ''' </summary>
        ''' <param name="ArqLog">Nome do arquivo, incluindo seu diretório.</param>
        ''' <param name="Msg">Mensagem a ser gravada.</param>
        ''' <remarks></remarks>
        Public Shared Sub GravaLog(ByVal ArqLog As String, ByVal Msg As String)
            For n As Integer = 1 To 10
                Try
                    Using log As New System.IO.StreamWriter(ArqLog, True)
                        log.WriteLine(Msg)
                        log.Close()
                    End Using
                    Exit Sub
                Catch
                End Try
                Threading.Thread.Sleep(10)
            Next
            Throw New Exception("[FALHA] ao tentar gravar em arquivo de log uma ocorrência.")
        End Sub

        ''' <summary>
        ''' Retorna URI raíz do site, considerando as variáveis de ambiente site_url e url_site.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function UriRaiz(Optional ByVal Compl As String = "") As String
            Static Uri As String = ""
            If Uri = "" Then
                Dim Ret As String = WebConf("site_url")
                If IsNothing(Ret) Then Ret = WebConf("url_site")
                If IsNothing(Ret) Then
                    Dim ctx As HttpContext = HttpContext.Current
                    Ret = Microsoft.VisualBasic.Left(ctx.Request.Url.AbsoluteUri, InStrRev(ctx.Request.Url.AbsoluteUri, ctx.Request.Url.LocalPath))
                End If
                Uri = Ret
            End If
            Return Uri
        End Function

        ''' <summary>
        ''' Verifica existência de item na lista sem quebra por delimitador.
        ''' </summary>
        ''' <param name="Lista">Lista de objetos onse a pesquisa ocorrerá.</param>
        ''' <param name="Conteudo">Conteúdo pesquisado.</param>
        ''' <param name="Atributo">Atributo opcional. Na falta deste, nome será utilizado.</param>
        ''' <returns>Retorna true caso encontre ou false caso contrário.</returns>
        ''' <remarks></remarks>
        Public Shared Function Exists(ByVal Lista As Object, ByVal Conteudo As String, Optional ByVal Atributo As String = "") As Boolean
            For Each Obj As Object In Lista
                If Prop(Obj, Atributo) = Conteudo Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Verifica existência de item texto na lista, quebrando texto origem conforme delimitador "UM;DOIS;TRES".
        ''' </summary>
        ''' <param name="Lista">Texto ou lista de objetos onde conteúdo será pesquisado.</param>
        ''' <param name="Conteudo">Conteúdo a ser pesquisado, que pode ser um texto ou um objeto.</param>
        ''' <param name="Delimit">Delimitador para o caso de origem como texto, que será decomposto de acordo com o delimitador.</param>
        ''' <param name="IgnoreCase">Opção de ignorar diferenciação de maiúsculas e minúsculas.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TemNaLista(ByVal Lista As Object, ByVal Conteudo As Object, Optional ByVal Delimit As String = ";", Optional ByVal IgnoreCase As Boolean = True) As Boolean
            Dim ListaObj As Object = Nothing
            If TypeOf Lista Is String Then
                ListaObj = Split(Lista, Delimit)
            ElseIf TypeOf Lista Is ArrayList OrElse TypeOf Lista Is Array Then
                ListaObj = Lista
            End If
            If Not IsNothing(Lista) Then
                For Each Elem As Object In ListaObj
                    If TypeOf Elem Is String Then
                        If Compare(Elem, Conteudo, IgnoreCase) Then
                            Return True
                        End If
                    Else
                        If Elem = Conteudo Then
                            Return True
                        End If
                    End If
                Next
                Return False
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Procura objeto através de um atributo.
        ''' </summary>
        ''' <param name="Lista">Lista de objetos na qual a pesquisa ocorrerá.</param>
        ''' <param name="Conteudo">Conteúdo a ser pesquisado.</param>
        ''' <param name="Atributo">Atributo considerado para pesquisa. Na sua ausência, nome será escolhido.</param>
        ''' <returns>Retorna o item encontrado ou nothing caso não haja coincidência de atributo.</returns>
        ''' <remarks></remarks>
        Shared Function ObjFindByAtt(ByVal Lista As Object, ByVal Conteudo As Object, Optional ByVal Atributo As String = "") As Object
            For Each Obj As Object In Lista
                If Prop(Obj, Atributo) = Conteudo Then
                    Return Obj
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Executa pausa na thread atual como sleep.
        ''' </summary>
        ''' <param name="Segundos">Segundos de pausa.</param>
        ''' <remarks></remarks>
        Shared Sub Espera(ByVal Segundos As Double)
#If _MYTYPE = "WindowsForms" Then
        Dim n As Date = Now
        Do While (Now - n).TotalSeconds < Segundos
            If Application.MessageLoop() Then
                Application.DoEvents()
            End If
        Loop
#Else
            System.Threading.Thread.Sleep(Segundos * 1000)
#End If
        End Sub

        ''' <summary>
        ''' Troca de texto enquanto este for encontrado na string de origem.
        ''' </summary>
        ''' <param name="Texto">Texto no qual ocorrerá a troca.</param>
        ''' <param name="De">Texto que será trocado. Enquanto este for encontrado, será substituído.</param>
        ''' <param name="Para">Texto a ser colocado no local do texto encontrado.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Shared Function ReplRepl(ByVal Texto As String, ByVal De As String, ByVal Para As String) As String
            Do While InStr(Texto, De) <> 0
                Texto = Replace(Texto, De, Para)
            Loop
            Return Texto
        End Function

        ''' <summary>
        ''' Procura node em uma árvore através de atributos previstos como enum NodeCampo.
        ''' </summary>
        ''' <param name="Arvore">Árvore onde ocorrerá a procura.</param>
        ''' <param name="Campo">Atributos no qual será baseada a procura.</param>
        ''' <param name="Conteudo">Conteúdo que deverá constar no atributo.</param>
        ''' <returns>Retorna o threenode, caso seja encontrado ou Nothing.</returns>
        ''' <remarks></remarks>
        Shared Function ProcuraNode(ByVal Arvore As TreeNodeCollection, ByVal Campo As NodeCampo, ByVal Conteudo As String) As TreeNode
            If Not IsNothing(Conteudo) Then
                For Each No As TreeNode In Arvore
                    Dim Ret As TreeNode = Nothing

                    If Campo = NodeCampo.NavigateUrl AndAlso TemNaLista(Conteudo, No.NavigateUrl) Then
                        Ret = No
                    ElseIf Campo = NodeCampo.Text AndAlso RegexGroup(No.Text, "(^|<div .*>)" & Conteudo & "($|</div>)").Success Then
                        Ret = No
                    ElseIf Campo = NodeCampo.TooolTip AndAlso Compare(No.ToolTip, Conteudo) Then
                        Ret = No
                    ElseIf Campo = NodeCampo.ValuePath AndAlso Compare(No.ValuePath, Conteudo) Then
                        Ret = No
                    Else
                        Ret = ProcuraNode(No.ChildNodes, Campo, Conteudo)
                    End If
                    If Not IsNothing(Ret) Then
                        Return Ret
                    End If
                Next
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Busca de node com base em conteúdo passado por variável GET.
        ''' </summary>
        ''' <param name="Arvore">Árvore a ser pesquisada.</param>
        ''' <param name="URL">URL tratada contendo a variável.</param>
        ''' <param name="SessaoVarNome">Nome da variável a ser extraída da URL.</param>
        ''' <param name="SessaoVarDef">Valor default da variável a ser considerada na falta da variável.</param>
        ''' <returns>Node encontrado ou nothing.</returns>
        ''' <remarks></remarks>
        Shared Function ProcuraNodeURLComplSessao(ByVal Arvore As TreeNodeCollection, ByVal URL As String, ByVal SessaoVarNome As String, ByVal SessaoVarDef As String) As TreeNode
            URL = URL & IIf(URL.IndexOf("?") <> -1, "&", "?") & SessaoVarNome & "=" & NZ(HttpContext.Current.Session(SessaoVarNome), SessaoVarDef)
            Return ProcuraNode(Arvore, NodeCampo.NavigateUrl, URL)
        End Function

        ''' <summary>
        ''' Obtém parte da string considerando parâmetro de início e final, como exemplo desta instrução existente no C++.
        ''' </summary>
        ''' <param name="Variavel">Variável texto a ser tratado.</param>
        ''' <param name="Inicio">Posição inicial a considerar, partindo do 0. Número negativo considera ponto a partir do fim do texto (-1 é o último caracter).</param>
        ''' <param name="Final">Posição final do texto a considerar partindo do 0. Número negativo considera ponto a partir do fim (-2 significa texto até o penúltimo caracter).</param>
        ''' <returns>Retorna parte do texto selecionado entre início e fim indicados.</returns>
        ''' <remarks></remarks>
        Shared Function StrStr(ByVal Variavel As String, ByVal Inicio As Integer, Optional ByVal Final As Integer = Nothing) As String
            If Inicio < 0 Then
                Inicio = (Len(Variavel) + Inicio)
            End If
            If Not NZ(Final, 0) = 0 Then
                If Final < 0 Then
                    Final = (Len(Variavel) + Final) - 1
                End If
                Return Variavel.Substring(Inicio, Final - Inicio + 1)
            End If
            Return Variavel.Substring(Inicio)
        End Function

        ''' <summary>
        ''' Concatena um conjunto de expressões separando-as ou não por um delimitador especificado.
        ''' </summary>
        ''' <param name="Delim">Um caractere ou expressão que será colocada entre as outras.</param>
        ''' <param name="DelimAlternativo">Um caractere ou expressão que será substituída por Delim.</param>
        ''' <param name="Inicial">Um caractere ou expressão que será colocada no início da string.</param>
        ''' <param name="Segmentos">O conjunto de expressões que será concatenado.</param>
        ''' <returns>Retorna uma string com todos os objetos de Segmentos concatenados.</returns>
        ''' <remarks></remarks>
        Shared Function ExprExpr(ByVal Delim As String, ByVal DelimAlternativo As String, ByVal Inicial As String, ByVal ParamArray Segmentos() As Object) As String
            For Each item As Object In Segmentos
                If Not IsNothing(item) Then
                    If Not TypeOf item Is String Then
                        Inicial &= ExprExpr(Delim, DelimAlternativo, Inicial, item)
                    End If
                    If Not IsNothing(DelimAlternativo) AndAlso DelimAlternativo <> "" Then
                        item = item.Replace(DelimAlternativo, Delim)
                    End If
                    If item <> "" Then
                        If Inicial <> "" Then
                            If Inicial.EndsWith(Delim) AndAlso item.StartsWith(Delim) Then
                                Inicial &= CType(item, String).Substring(Delim.Length)
                            ElseIf Inicial.EndsWith(Delim) OrElse item.StartsWith(Delim) Then
                                Inicial &= item
                            Else
                                Inicial &= Delim & item
                            End If
                        Else
                            Inicial &= item
                        End If
                    End If
                End If
            Next
            Return Inicial
        End Function


        ''' <summary>
        ''' Obtém path de arquivo concatenando partes de diretório e arquivo, como por exemplo: "c:\" + "dir1" + "subdir1" + "arquivo.txt". A rotina colocará as barras. Pode-se passar ~/ também.
        ''' </summary>
        ''' <param name="Segmentos">Segmentos a serem concatenados em forma de diretório e arquivo.</param>
        ''' <returns>Retorna path resultante da concatenação de segmentos, evitando barras repetidas.</returns>
        ''' <remarks></remarks>
        Shared Function FileExpr(ByVal ParamArray Segmentos() As String) As String
            Dim Raiz As String = New System.Web.UI.Control().ResolveUrl("~/").Replace("/", "\")
            Dim Arq As String = ExprExpr("\", "/", "", Segmentos)
            If Arq.StartsWith(Raiz) Then
                Arq = "~\" & Mid(Arq, Len(Raiz) + 1)
            End If

            If Arq.StartsWith("~\") Then
                Arq = HttpContext.Current.Server.MapPath(Arq)
            End If
            Return Arq
        End Function


        ''' <summary>
        ''' Contatena URL evitando barras repetidas.
        ''' </summary>
        ''' <param name="Segmentos">São os trechos a serem concatenados, podendo ser mais de dois.</param>
        ''' <returns>Retorna expressão de segmentos concatenados.</returns>
        ''' <remarks></remarks>
        Shared Function URLExpr(ByVal ParamArray Segmentos() As Object) As String
            Dim URL As String = ExprExpr("/", "\", "", Segmentos)
            If Regex.Match(URL, "(?is)^[a-z]:/").Success Then
                URL = URL.Replace(HttpContext.Current.Server.MapPath("~/").Replace("\", "/"), "~/")
            End If
            Return URL
        End Function

        ''' <summary>
        ''' Concatena segmentos incluindo separador entre os termos existentes.
        ''' </summary>
        ''' <param name="Separador">Termo a ser colocado entre os segmentos.</param>
        ''' <param name="Segmentos">Itens em paramarray a ser concatenado com o separador.</param>
        ''' <returns>Retorna texto resultante da concatenação dos segmentos utilizando separador.</returns>
        ''' <remarks></remarks>
        Shared Function SegmExpr(ByVal Separador As String, ByVal ParamArray Segmentos() As Object) As String
            Dim Ret As String = ""
            For Each Item As String In Segmentos
                If Item <> "" Then
                    If Ret <> "" AndAlso Not Ret.EndsWith(Separador) Then
                        Ret &= Separador
                    End If
                    Ret &= Item
                End If
            Next
            Return Ret
        End Function

        ''' <summary>
        ''' Cancela cache da página indicada.
        ''' </summary>
        ''' <param name="Pagina">Página que deverá ter o cache cancelado.</param>
        ''' <remarks></remarks>
        Shared Sub NoCache(ByVal Pagina As Page, Optional ByVal Tipo As NoCacheTipo = NoCacheTipo.PaginaExpirada)
            If Tipo = NoCacheTipo.PaginaExpirada Then
                Pagina.Response.Cache.SetExpires(DateTime.Now)
            ElseIf Tipo = NoCacheTipo.SemHistorico Then
                Pagina.Response.Cache.SetExpires(DateTime.Now)
                Pagina.Response.Cache.SetNoStore()
            End If
            Pagina.Response.AppendHeader("pragma", "no-cache")
        End Sub

        ''' <summary>
        ''' Busca nos subcontrols de um objeto, itens do tipo especificado.
        ''' </summary>
        ''' <param name="Obj">Container a ser pesquisado.</param>
        ''' <param name="Prefix">Prefixo dos controles considerados.</param>
        ''' <param name="Tipo">Tipo no formato texto ex:System.String.</param>
        ''' <returns>Retorna lista de objetos encontrados, que atendam o critério.</returns>
        ''' <remarks></remarks>
        Shared Function ItemsDoTipo(ByVal Obj As Object, ByVal Prefix As String, ByVal Tipo As String) As List(Of Object)
            Dim Lista As List(Of Object) = New List(Of Object)
            For Each Item As Object In Form.Controles(Obj, Prefix)
                If Compare(Item.GetType.ToString, Tipo) Then
                    Lista.Add(Item)
                End If
            Next
            Return Lista
        End Function

        ''' <summary>
        ''' Adiciona itens de uma origem para uma coleção destino.
        ''' </summary>
        ''' <param name="Destino">A coleção na qual serão adicionados os itens.</param>
        ''' <param name="Origem">Uma coleção contendo os itens a serem enumerados (for each) para cópia.</param>
        ''' <returns>Retorna a quantidade de itens copiados.</returns>
        ''' <remarks></remarks>
        Shared Function CopiaItens(ByRef Destino As Object, ByRef Origem As Object) As Integer
            Dim Qtd As Integer = 0
            For Each Item As Object In Origem
                Destino.Add(Item)
                Qtd += 1
            Next
            Return Qtd
        End Function

        ''' <summary>
        ''' Mediante um conteúdo, apresenta ou não paineis.
        ''' </summary>
        ''' <param name="Container">Container onde ocorrerá a busca dos controles.</param>
        ''' <param name="Prefixo">Prefixo dos paineis que serão ocultados.</param>
        ''' <param name="Escolha">Sufixo do painel que será apresentado.</param>
        ''' <remarks></remarks>
        Shared Sub SelecionaDivisaoPainel(ByVal Container As Object, ByVal Prefixo As String, ByVal Escolha As String)
            If Escolha.StartsWith("[") And Escolha.EndsWith("]") Then
                Escolha = StrStr(Escolha, 1, -1)
            End If

            For Each Ctl As Control In Form.Controles(Container, Prefixo)
                Dim item As String = Mid(Prop(Ctl, "ID"), Len(Prefixo) + 1)
                If item <> "" Then
                    Ctl.Visible = (item = Escolha)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Retorna número aleatório para ser utilizado como arquivo temporário.
        ''' </summary>
        ''' <param name="Dir">Diretório onde será criado o arquivo. Vazio para obter o diretório default configurado em web.config.</param>
        ''' <returns>Retorna diretório e arquivo temporário.</returns>
        ''' <remarks></remarks>
        Shared Function TemporaryFile(Optional ByVal Dir As String = "", Optional ByVal Extensao As String = "tmp") As String
            Dim DirArq As String = ""
            If Dir = "" Then
                Dir = TemporaryDir()
            End If
            Dim Arq As String = ""
            Dim Vezes As Integer = 0
            Do While Arq = ""
                For z As Integer = 0 To 12
                    Arq &= Int(Rnd(Now.Millisecond) * 10)
                Next
                Arq &= "." & Extensao
                DirArq = FileExpr(Dir, Arq)
                If System.IO.File.Exists(DirArq) Then
                    Arq = ""
                End If
                Vezes += 1
                If Vezes > 500 Then
                    Throw New Exception("Tentativa de busca de arquivo temporário falho (máximo de 500 tentativas atingido).")
                    Exit Function
                End If
            Loop
            Return DirArq
        End Function

        ''' <summary>
        ''' Insere tabuladores no início de cada linha.
        ''' </summary>
        ''' <param name="Texto">Texto a ser tratado.</param>
        ''' <param name="Tabulador">Texto que será utilizado como tabulador (default são quatro espaços).</param>
        ''' <param name="QuebradeLinha">Marcador de final de linha. Default é vbcrlf.</param>
        ''' <returns>Retorna texto tratado.</returns>
        ''' <remarks></remarks>
        Shared Function InsereTab(ByVal Texto As String, Optional ByVal Tabulador As String = "    ", Optional ByVal QuebradeLinha As String = vbCrLf) As String
            If Texto = "" Then
                Return Texto
            End If
            Return Tabulador & Replace(Texto, QuebradeLinha, QuebradeLinha & Tabulador)
        End Function

        ''' <summary>
        ''' Mostra uma mensagem alerta Javascript e realiza o redirecionamento cliente para uma URL especificada
        ''' </summary>
        ''' <param name="objPage">Objeto página aspx</param>
        ''' <param name="Mensagem">Mensagem a ser mostrada</param>
        ''' <param name="URL">URL para onde o cliente será redirecionado após apresentação da mensagem</param>
        ''' <remarks></remarks>
        Public Shared Sub ShowJSMessage(ByRef ObjPage As Page, ByVal Mensagem As String, Optional ByVal URL As String = "")
            Dim conteudo As New StringBuilder
            conteudo.Append("<script>")
            conteudo.Append("   alert(""" & Mensagem.Replace("""", "").Replace(vbCrLf, "\n").Replace(Chr(10), "\n") & """);")
            If URL = "" Then
                conteudo.Append("   setTimeout('__doPostBack()', 0);")
            Else
                conteudo.Append("   window.location.href = """ & ObjPage.ResolveUrl(URL) & """;")
            End If
            conteudo.Append("</script>")
            ObjPage.ClientScript.RegisterClientScriptBlock(ObjPage.GetType(), "ShowJSMessageAndRedirect", conteudo.ToString)
        End Sub

        ''' <summary>
        ''' Rotina que cria texto para submissão de form contendo evento em ASP.NET.
        ''' </summary>
        ''' <param name="ObjPage">Página que receberá o evento.</param>
        ''' <param name="Alvo">Alvo.</param>
        ''' <param name="Argumento">Argumento.</param>
        ''' <remarks></remarks>
        Public Shared Sub ExecPostBack(ByVal ObjPage As Page, Optional ByVal Alvo As String = "", Optional ByVal Argumento As String = "")
            Dim conteudo As New StringBuilder
            conteudo.Append("<script>")
            conteudo.Append("   setTimeout('__doPostBack(""" & Alvo & """, """ & Argumento & """)', 0);")
            conteudo.Append("</script>")
            ObjPage.ClientScript.RegisterClientScriptBlock(ObjPage.GetType(), "ShowJSMessageAndRedirect", conteudo.ToString)
        End Sub


        ''' <summary>
        ''' formato do match de troca para link [link:descrição|url]
        ''' link:url|descrição 
        ''' img:url|legenda
        ''' imgbut:urlimagem|urllink|descrição
        ''' arquivo:url:descrição [[[[continuar]]]][[[[continuar]]]][[[[continuar]]]]
        ''' arvore...
        ''' tabela(col,lin):itens
        ''' 
        ''' classe iniciada. [[parei por causa da pressa para fazer sepon]]
        ''' </summary>
        ''' <remarks></remarks>
        Class Exibe
            Private _page As Page
            Sub New(ByVal Page As Page)
                _page = Page
            End Sub

            Public Function HTMLReplImgBut(ByVal m As Match) As String
                Return "\\{<a href=""" & _page.ResolveUrl(m.Groups(2).Value) & """><img src=""" & ImageURL(_page, m.Groups(1).Value) & """ alt=""" & HttpUtility.HtmlEncode(m.Groups(3).Value) & """/></a>\\}"
            End Function

            Public Function HTMLReplLink(ByVal m As Match) As String
                Return "\\{<a href=""" & _page.ResolveUrl(m.Groups(1).Value) & """>" & HttpUtility.HtmlEncode(m.Groups(2).Value) & "</a>\\}"
            End Function

            Public Function HTMLReplImg(ByVal m As Match) As String
                Return "\\{<img src=""" & _page.ResolveUrl(m.Groups(1).Value) & """ alt=""" & HttpUtility.HtmlEncode(m.Groups(2).Value) & """ />\\}"
            End Function
        End Class

        ''' <summary>
        ''' Obtém a URL de um arquivo de imagem de acordo com o diretório especificado.
        ''' </summary>
        ''' <param name="Page">Página que precisa da URL.</param>
        ''' <param name="Arquivo">O nome do arquivo de imagem.</param>
        ''' <param name="Diretorio">Vazio para o diretório padrão. public para o diretório de imagens públicas e priv para o diretório de imagens privadas.</param>
        ''' <returns>Retorna a URL do arquivo de imagem requerido de acordo com o diretório especificado.</returns>
        ''' <remarks></remarks>
        Public Shared Function ImageURL(ByVal Page As Page, ByVal Arquivo As String, Optional ByVal Diretorio As String = "") As String
            Diretorio = NZ(Diretorio, "")

            Select Case LCase(Diretorio)
                Case ""
                    Diretorio = "~/img/"
                Case "public"
                    Diretorio = "~/img_public/"
                Case "priv"
                    Diretorio = URLExpr("~/img_priv", Icraft.IcftBase.Logon(Page).Usuario, "/")
                Case Else
                    Diretorio = "~/img_" & Diretorio & "/"
            End Select
            Return Page.ResolveUrl(URLExpr(Diretorio, Arquivo))
        End Function

        ''' <summary>
        ''' Obtém o endereço físico do arquivo de imagem especificado.
        ''' </summary>
        ''' <param name="Page">Página que precisa do endereço da imagem.</param>
        ''' <param name="Arquivo">O nome do arquivo de imagem.</param>
        ''' <param name="Diretorio">Vazio para o diretório padrão. public para o diretório de imagens públicas e priv para o diretório de imagens privadas.</param>
        ''' <returns>Retorna o endereço físico do arquivo de imagem requerido de acordo com o diretório especificado.</returns>
        ''' <remarks></remarks>
        Public Shared Function ImageArq(ByVal Page As Page, ByVal Arquivo As String, Optional ByVal Diretorio As String = "") As String
            Return Page.MapPath(ImageURL(Page, Arquivo, Diretorio))
        End Function

        ''' <summary>
        ''' Obtém texto html codificado transformando quebras de linha em parágrafos.
        ''' </summary>
        ''' <param name="Page">Página que pediu o texto codificado.</param>
        ''' <param name="Texto">Texto que será codificado.</param>
        ''' <returns>Retorna a string passada em Texto com codificação HTML.</returns>
        ''' <remarks></remarks>
        Public Shared Function ExibeHTMLEnc(ByVal Page As Page, ByVal Texto As String) As String
            Dim Ret As New StringBuilder

            For Each L1 As String In Regex.Split(Texto, "(\\\\{[^}]*\\\\})")
                If L1.StartsWith("\\{") And L1.EndsWith("\\}") Then
                    Ret.Append(StrStr(L1, 3, -3))
                Else
                    For Each L2 As String In Split(L1, vbCrLf)
                        Ret.AppendLine("<p>")
                        Ret.Append("    ")
                        Ret.AppendLine(HttpUtility.HtmlEncode(L2))
                        Ret.AppendLine("</p>")
                    Next
                End If
            Next
            Return Ret.ToString
        End Function

        ''' <summary>
        ''' Transforma texto comum em texto html.
        ''' </summary>
        ''' <param name="Page">Objeto page que está requisitando o texto.</param>
        ''' <param name="Texto">Texto que será transformado em html.</param>
        ''' <returns>Retorna o texto transformado em html ou o próprio texto caso este esteja entre &lt;html&gt; e &lt;/html&gt;</returns>
        ''' <remarks></remarks>
        Public Shared Function ExibeHTML(ByVal Page As Page, ByVal Texto As String) As String
            ' Dim Tag As String = Icraft.IcftBase.RegexMascTags("p", 4).Replace("p", "p|div|h1|strong")
            Dim Tag As String = Icraft.IcftBase.RegexMascTags(".+", 4)
            ' "(?is)\<html\>.*\<\/html\>"
            If Regex.Match(Texto, Tag).Success Then
                Return Texto
            End If

            Dim ex As New Exibe(Page)
            Dim mev As New MatchEvaluator(AddressOf ex.HTMLReplLink)
            Texto = Regex.Replace(Texto, "\[link:(.*)\|(.*)\]", mev)

            mev = New MatchEvaluator(AddressOf ex.HTMLReplImg)
            Texto = Regex.Replace(Texto, "\[img:(.*)\|(.*)\]", mev)

            mev = New MatchEvaluator(AddressOf ex.HTMLReplImgBut)
            Texto = Regex.Replace(Texto, "\[imgbut:(.*)\|(.*)\|(.*)\]", mev)

            Texto = Regex.Replace(Texto, "//{([^}]*)//}", "$1")

            Return ExibeHTMLEnc(Page, Texto)
        End Function

        Enum ExibeSegsOpc
            xh_ymin_zseg
            hh_mm_ss
            x_horas_y_minutos_e_z_segundos
            z_segundos
            hh_mm
        End Enum
        Public Shared Function ExibeSegs(ByVal QtdSegundos As Integer, ByVal Opc As ExibeSegsOpc) As String
            Dim Segs As Integer = QtdSegundos
            Dim Horas As Integer = Int(Segs / 3600)
            Segs -= Horas * 3600
            Dim Mins As Integer = Int(Segs / 60)
            Segs -= Mins * 60

            Select Case Opc
                Case ExibeSegsOpc.hh_mm_ss
                    Return Format(Horas, "00") & ":" & Format(Mins, "00") & ":" & Format(Segs, "00")
                Case ExibeSegsOpc.x_horas_y_minutos_e_z_segundos
                    Return Horas & Horas & Pl(Horas, " Hora") & ", " & Mins & Pl(Mins, " Minuto") & " e " & Segs & Segs & Pl(Segs, " Segundo")
                Case ExibeSegsOpc.xh_ymin_zseg
                    Return Horas & "h " & Mins & "min " & Segs & "seg"
                Case ExibeSegsOpc.hh_mm
                    Return Format(Horas, "00") & ":" & Format(Mins, "00")
            End Select
            Return QtdSegundos & Pl(QtdSegundos, " Segundo")
        End Function



        ''' <summary>
        ''' Formatos possíveis de data.
        ''' </summary>
        ''' <remarks></remarks>
        Enum ExibeDataOpc
            p
            dd_de_mmmm_de_yyyy
            i
            mmmm_dth_yyyy
            a
            dd_mmm_yyyy
            ai
            dd_mmm_yyyy_i
            mmm_dd_yyyy
            mmm_dd_yyyy_i
            mmm_dd_yyyy_c
            mmmm_yyyy
            mmmm_yyyy_i
            mmmm_yyyy_c
            mmm
            mmm_i
            mmm_c
        End Enum

        ''' <summary>
        ''' Formata data para exibição correta
        ''' </summary>
        ''' <param name="Momento">Data para ser formatada.</param>
        ''' <param name="Opc">ExibeDataOpc que será utilizado na apresentação.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ExibeData(ByVal Momento As Object, ByVal Opc As ExibeDataOpc) As String
            Return ExibeData(Momento, Replace(Opc.ToString, "_", " "))
        End Function


        ''' <summary>
        ''' Formata data para exibição correta
        ''' </summary>
        ''' <param name="Momento">Data para ser formatada.</param>
        ''' <param name="Formato">Formato a ser apresentado. Padrão é dd/MM/yyyy HH:mm:ss.</param>
        ''' <returns>Retorna a data formatada.</returns>
        ''' <remarks></remarks>
        Public Shared Function ExibeData(ByVal Momento As Object, Optional ByVal Formato As String = "") As String
            Try
                Momento = CType(Momento, Date)
            Catch
                Momento = CDate(Nothing)
            End Try

            If Momento = CDate(Nothing) Then
                Return ""
            End If

            If Formato = "" Then
                Return Format(Momento, "dd/MM/yyyy HH:mm:ss")
            End If

            Dim MMP() As String = {"Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"}
            Dim MMI() As String = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
            Dim MMC() As String = {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"}


            Dim DIA As Integer = Microsoft.VisualBasic.Day(Momento)
            Dim MES As Integer = Month(Momento) - 1
            Dim ANO As Integer = Year(Momento)
            Dim SUF As Integer = DIA Mod 10
            Dim IDI As String = ""

            Select Case LCase(Formato)
                Case "p", "dd de mmmm de yyyy"
                    Return DIA & " de " & MMP(MES) & " de " & ANO
                Case "c", "dd de mmmm de yyyy c"
                    Return DIA & " de " & MMC(MES) & " de " & ANO
                Case "i", "mmmm dth, yyyy", "mmmm dth yyyy"
                    Return MMI(MES) & " " & DIA & Microsoft.VisualBasic.Switch(DIA > 10 And DIA < 14, "th", SUF = 1, "st", SUF = 2, "nd", SUF = 3, "rd", True, "th") & ", " & Format(ANO, "YYYY")
                Case "a", "dd mmm yyyy"
                    Return MMI(MES) & " " & DIA & Microsoft.VisualBasic.Switch(DIA > 10 And DIA < 14, "th", SUF = 1, "st", SUF = 2, "nd", SUF = 3, "rd", True, "th") & ", " & Format(ANO, "YYYY")
                Case "ai", "dd mmm yyyy i"
                    Return Format(DIA, "00") & " " & UCase(Microsoft.VisualBasic.Left(MMI(MES), 3)) & " " & Format(ANO, "0000")
                Case "mmm dd, yyyy", "mmm dd yyyy"
                    Return UCase(Microsoft.VisualBasic.Left(MMP(MES), 3)) & " " & DIA & ", " & Format(ANO, "0000")
                Case "mmm dd, yyyy i", "mmm dd yyyy i"
                    Return UCase(Microsoft.VisualBasic.Left(MMI(MES), 3)) & " " & DIA & ", " & Format(ANO, "0000")
                Case "mmm dd, yyyy c", "mmm dd yyyy c"
                    Return UCase(Microsoft.VisualBasic.Left(MMC(MES), 3)) & " " & DIA & ", " & Format(ANO, "0000")
                Case "mmmm, yyyy", "mmmm yyyy"
                    Return MMP(MES) & ", " & Format(ANO, "0000")
                Case "mmmm/yy", "mmmm yy"
                    Return MMP(MES) & "/" & Format(Momento, "yy")
                Case "mmmm, yyyy i", "mmmm yyyy i"
                    Return MMI(MES) & ", " & Format(ANO, "0000")
                Case "mmmm, yyyy c", "mmmm yyyy c"
                    Return MMC(MES) & ", " & Format(ANO, "0000")
                Case "mmm"
                    Return UCase(Microsoft.VisualBasic.Left(MMP(MES), 3))
                Case "mmm i"
                    Return UCase(Microsoft.VisualBasic.Left(MMI(MES), 3))
                Case "mmm c"
                    Return UCase(Microsoft.VisualBasic.Left(MMC(MES), 3))
            End Select

            Return Format(Momento, Formato)
        End Function

        ''' <summary>
        ''' Formata data para gravação correta de acordo com o banco de dados
        ''' </summary>
        ''' <param name="data">Data para ser formatada.</param>
        ''' <param name="banco">Nome do banco de destino da data.</param>
        ''' <returns>Retorna a data formatada.</returns>
        ''' <remarks></remarks>
        Public Shared Function GravaData(ByVal data As Date, ByVal banco As TipoBaseSQL) As String
            Dim data_destino As String = Now
            Select Case banco
                Case TipoBaseSQL.MSAccess Or TipoBaseSQL.Oracle
                    data_destino = data.Year.ToString("0000") & "/" & data.Month.ToString("00") & "/" & data.Day.ToString("00")
                Case TipoBaseSQL.MySQL
                    data_destino = data.Year.ToString("0000") & "-" & data.Month.ToString("00") & "-" & data.Day.ToString("00")
            End Select
            Return data_destino
        End Function

        ''' <summary>
        ''' Retorna o primeiro item do array procurado pelo objeto ou atributo.
        ''' </summary>
        ''' <param name="Lista">Array a ser pesquisado.</param>
        ''' <param name="Conteudo">Conteúdo que será procurado ou no índice do array ou em algum atributo.</param>
        ''' <param name="Atributo">Vazio para procurar na posição ou nome para pesquisa pela propriedade attribute.</param>
        ''' <param name="Inicio">Zero para procurar do início ou posição inicial do array.</param>
        ''' <returns>Retorna o item do array encontrado.</returns>
        ''' <remarks></remarks>
        Shared Function ArrayFindByAtt(ByVal Lista As Array, ByVal Conteudo As Object, Optional ByVal Atributo As String = "", Optional ByVal Inicio As Integer = 0) As Object
            Dim pos As Integer = ArrayIndexFindByAtt(Lista, Conteudo, Atributo, Inicio)
            If pos = -1 Then
                Return Nothing
            End If
            Return Lista(pos)
        End Function

        ''' <summary>
        ''' Retorna posição do primeiro item no array pelo objeto ou atributo.
        ''' </summary>
        ''' <param name="Lista">Array a ser pesquisado.</param>
        ''' <param name="Conteudo">Conteúdo que será procurado ou no índice do array ou em algum atributo.</param>
        ''' <param name="Atributo">Vazio para procurar na posição ou nome para pesquisa pela propriedade attribute.</param>
        ''' <param name="Inicio">Zero para procurar do início ou posição inicial do array.</param>
        ''' <returns>Retorna posição do item do array encontrado.</returns>
        ''' <remarks></remarks>
        Shared Function ArrayIndexFindByAtt(ByVal Lista As Array, ByVal Conteudo As String, Optional ByVal Atributo As String = "", Optional ByVal Inicio As Integer = 0) As Integer
            Dim z As Integer, item As Object = Nothing
            For z = 0 To Lista.Length - 1
                If Atributo = "" Then
                    item = Lista(z)
                Else
                    item = Lista(z).Attributes(Atributo)
                End If
                If Compare(item, Conteudo) Then
                    Exit For
                End If
            Next
            If z >= Lista.Length Then
                Return -1
            End If
            Return z
        End Function

        ''' <summary>
        ''' Carrega qualquer script no corpo da página.
        ''' </summary>
        ''' <param name="Pag">Página onde será carregado o script.</param>
        ''' <param name="NomeScript">Nome que o script terá no corpo.</param>
        ''' <param name="SegScript">Pode ser um nome de arquivo '~/dir/arquivo.js' ou um bloco '&lt;script&gt;....&lt;/script&gt;'.</param>
        ''' <remarks></remarks>
        Shared Sub IncluiScript(ByRef Pag As Object, ByVal NomeScript As String, Optional ByVal SegScript As String = "")
            If SegScript = "" Then
                SegScript = Pag.resolveUrl("~\inc\" & NomeScript)
            End If
            Dim cs As ClientScriptManager = Pag.ClientScript
            If (Not cs.IsClientScriptIncludeRegistered(Pag.GetType(), NomeScript)) Then
                If SegScript.IndexOf("<script") <> -1 Then
                    cs.RegisterClientScriptBlock(Pag.GetType(), NomeScript, SegScript)
                Else
                    cs.RegisterClientScriptInclude(Pag.GetType(), NomeScript, SegScript)
                End If
            End If
        End Sub

        ''' <summary>
        ''' Compara dois parâmetros com base em critério específicos para cada tipo.
        ''' </summary>
        ''' <param name="Param1">Primeiro parâmetro.</param>
        ''' <param name="Param2">Segundo parâmetro.</param>
        ''' <param name="IgnoreCase">Para ignorar diferença entre maiúsculo e minúsculo em comparações de strings.</param>
        ''' <returns>Retorna verdadeiro caso os itens sejam considerados iguais ou o contrário.</returns>
        ''' <remarks></remarks>
        Shared Function Compare(ByVal Param1 As Object, ByVal Param2 As Object, Optional ByVal IgnoreCase As Boolean = True) As Boolean
            If IsNothing(Param1) And IsNothing(Param2) Then
                Return True
            ElseIf IsNothing(Param1) Or IsNothing(Param2) Then
                Return False
            Else
                If Param1.GetType.ToString = Param2.GetType.ToString Then
                    If Param1.GetType.ToString = "System.String" Then
                        Return String.Compare(Param1, Param2, IgnoreCase) = 0
                    Else
                        Err.Raise(20000, "IcraftBase", "Compare com tipo não previsto " & Param1.GetType.ToString & ".")
                    End If
                End If
            End If
            Return False
        End Function

        ''' <summary>
        ''' Rotina que é executada quando qualquer CONTROLE é atualizado. Responsável pela atualização de outros CONTROLES dependentes.
        ''' </summary>
        ''' <param name="Controle">Controle que foi atualizado.</param>
        ''' <param name="e">Argumento padrão do sistema.</param>
        ''' <remarks></remarks>
        Shared Sub AtualizouControle(ByVal Controle As Object, Optional ByVal e As System.EventArgs = Nothing)
            If PropE(Controle, "Atualizar") <> "" Then
                For Each Campo As String In Split(PropE(Controle, "Atualizar"), ";")
                    Dim CampoRel As Object = Form.FindGeral(Controle.Parent, Campo)
                    If Not IsNothing(CampoRel) Then
                        If TypeOf CampoRel Is DropDownList Then
                            CarregaCombo(CampoRel)
                        Else
                            CampoRel.DataBind()
                        End If
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' Obtém diretório temporário, que corresponde ao param dir_temp do web.config.
        ''' </summary>
        ''' <returns>Retorna diretório para arquivos temporários sem a barra no final (ex: c:\inetpub\temp).</returns>
        ''' <remarks></remarks>
        Shared Function TemporaryDir(Optional ByVal Dir As String = "") As String
            If Dir <> "" Then
                Dim DirCompl As String = ""
                Dim Vezes As Integer = 0
                Do While DirCompl = ""
                    For z As Integer = 0 To 12
                        DirCompl &= Int(Rnd(Now.Millisecond) * 10)
                    Next
                    DirCompl = FileExpr(Dir, DirCompl)
                    If System.IO.Directory.Exists(DirCompl) Then
                        DirCompl = ""
                    End If
                    Vezes += 1
                    If Vezes > 100 Then
                        Throw New Exception("Tentativa de busca de diretório temporário falhou (máximo de 100 tentativas atingido).")
                        Exit Function
                    End If
                Loop
                Return DirCompl
            End If
            Return WebConf("dir_temp")
        End Function

        ''' <summary>
        ''' Retorna parâmetro específico do webconfig > appsetings.
        ''' </summary>
        ''' <param name="param">Identificação do connectionstring desejado.</param>
        ''' <returns>Objeto connectionstringsettings obtido a partir do configurationmanager.</returns>
        ''' <remarks></remarks>
        Shared Function WebConf(ByVal param As String) As String
            Return System.Configuration.ConfigurationManager.AppSettings(param)
        End Function

        ''' <summary>
        ''' Procura um controle em um determinado panel. Criada por CTL.FINDCONTROL não conseguir encontrar obj em PANEL.
        ''' </summary>
        ''' <param name="Origem">Objeto onde será feita a procura.</param>
        ''' <param name="Controle">Nome do controle a ser procurado.</param>
        ''' <returns>Retorna controle ou nothing caso não encontre. A procura é feita somente naquele nível. Utilize FORM.FINDCONTROL para encontrar controles entre os filhos e FORM.FINDGERAL para procurar entre filhos e pais.</returns>
        ''' <remarks></remarks>
        Shared Function FindControlEspecial(ByVal Origem As Object, ByVal Controle As String) As Object
            ' procura somente no primeiro nível
            ' rotina igual ao findcontrol original, mas resolvendo a questão de busca em painel
            ' >> painel tinha o conteúdo e findcontrol não retornava de forma alguma!!
            Dim ctl As Control = Nothing
            If TypeOf Origem Is Panel Then

                ' busca direta e caso não funcione, busca por item
                ctl = CType(Origem, Panel).FindControl(Controle)
                If Not IsNothing(ctl) Then
                    Return ctl
                End If
                Return FindControlEspecial(CType(Origem, Panel).Controls, Controle)
            ElseIf TypeOf Origem Is Web.UI.ControlCollection Or TypeOf Origem Is Windows.Forms.Form.ControlCollection Then

                ' aproveitei e implementei procura na coleção também
                For Each SubCtl As Control In Origem
                    If Compare(Prop(SubCtl, "ID"), Controle) Then
                        Return SubCtl
                    End If
                Next
                Return Nothing
            End If

            ' para qualquer outro objeto, faz busca com findcontrol
            ' caso ocorra erro, retorna nothing mesmo...
            Try
                ctl = Origem.findcontrol(Controle)
            Catch
            End Try
            Return ctl
        End Function

        ''' <summary>
        ''' Carrega qualquer script no corpo da página.
        ''' </summary>
        ''' <param name="Pag">Página onde será carregado o script.</param>
        ''' <param name="href">Caminho do css Ex:(~/diretorio/estilo.css)</param>    
        ''' <remarks></remarks>
        Shared Sub IncluiStyleSheet(ByVal Pag As Page, ByVal id As String, Optional ByVal href As String = "")
            If IsNothing(Pag.Header.FindControl(id)) Then
                If href = "" Then
                    href = Pag.ResolveUrl("~\inc\" & id)
                End If

                Dim link As New System.Web.UI.HtmlControls.HtmlLink
                link.ID = id
                link.Href = href
                link.Attributes("rel") = "stylesheet"
                link.Attributes("type") = "text/css"
                Pag.Header.Controls.Add(link)
            End If
        End Sub

        ''' <summary>
        ''' Escrever
        ''' </summary>
        ''' <param name="Texto"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Shared Function AbstrCarac(ByVal Texto As String) As String
            Dim result As Integer = 0
            Dim ll As Integer = Len(Texto)
            For z = 0 To ll - 1
                For z1 = 0 To ll - 1
                    result += Asc(Texto.Substring(z1, 1)) + Asc(Texto.Substring(ll - z1 - 1, 1)) + ll
                Next
                Mid(Texto, ll - z, 1) = Base36Alga(result Mod 36)
            Next
            Return Texto
        End Function

        ''' <summary>
        ''' Transforma um número qualquer em caracteres utilizando Base36.
        ''' </summary>
        ''' <param name="NUM">O número que será transformado.</param>
        ''' <param name="NumCasas">A quantidade de casas que a string de retorno conterá. O número tem precedência sobre o resultado.</param>
        ''' <returns>Retorna uma string com o tamanho de NumCasas contendo o número passado em NUM convertido através de Base36.</returns>
        ''' <remarks></remarks>
        Shared Function Base36(ByVal NUM As Integer, Optional ByVal NumCasas As Integer = 0) As String
            Dim TOT As Integer = NUM
            Dim RET As String = ""
            Do While TOT > 0 And (NumCasas = 0 OrElse Len(RET) < NumCasas)
                Dim REST As Integer = TOT Mod 36
                RET = Base36Alga(REST) & RET
                TOT = Int(TOT / 36)
            Loop
            For z = Len(RET) To NumCasas - 1
                RET = "0" & RET
            Next
            Return RET
        End Function

        ''' <summary>
        ''' Transforma um número entre 0 e 35 em um caractere utilizando Base36.
        ''' </summary>
        ''' <param name="IDP">O número que será transformado.</param>
        ''' <returns>Retorna o caractere para o qual o número foi transformado.</returns>
        ''' <remarks></remarks>
        Shared Function Base36Alga(ByVal IDP As Integer) As String
            If IDP < 0 Then
                Return "0"
            ElseIf IDP <= 9 Then
                Return Chr(IDP + Asc("0"))
            ElseIf IDP <= 34 Then
                Return Chr((IDP - 10) + Asc("A"))
            End If
            Return "Z"
        End Function

        ''' <summary>
        ''' Apaga arquivos temporários ignorando erros.
        ''' </summary>
        ''' <param name="Tmps">Arraylist contendo arquivos temporários.</param>
        ''' <remarks></remarks>
        Shared Sub ApagaTemps(ByVal Tmps As ArrayList)
            If Not IsNothing(Tmps) Then
                For Each tmp As String In Tmps
                    Try
                        System.IO.File.Delete(tmp)
                    Catch
                    End Try
                Next
            End If
        End Sub

        ''' <summary>
        ''' Retorna lista de arquivos contidos no diretório
        ''' </summary>
        ''' <param name="Diretorio">Caminho para disco onde será feita a pesquisa.</param>
        ''' <returns>Arraylist contendo os arquivos pesquisados.</returns>
        ''' <remarks></remarks>
        Public Shared Function ListaDir(ByVal Diretorio As String, Optional ByVal Criterio As String = "*.*") As ArrayList
            Dim ret As New ArrayList
            Diretorio = FileExpr(Diretorio)
            For Each fl As String In System.IO.Directory.GetFiles(Diretorio, Criterio)
                ret.Add(fl)
            Next
            For Each dr As String In System.IO.Directory.GetDirectories(Diretorio)
                ret.AddRange(ListaDir(dr, Criterio))
            Next
            Return ret
        End Function


        ''' <summary>
        ''' Retorna connectionstring específico do webconfig > connectionstring.
        ''' </summary>
        ''' <param name="param">Identificação do connectionstring desejado.</param>
        ''' <returns>Objeto connectionstringsettings obtido a partir do configurationmanager.</returns>
        ''' <remarks></remarks>
        Shared Function WebConn(ByVal Param As String) As System.Configuration.ConnectionStringSettings
#If _MyType = "WindowsForms" Then
        Return System.Configuration.ConfigurationManager.ConnectionStrings(Param)
#Else
            Return System.Configuration.ConfigurationManager.ConnectionStrings(Param)
#End If
        End Function

        ''' <summary>
        ''' Retorna um connectionstring a partir da informação de um connectionstring ou string indicativa da conexão no WebConfig.
        ''' </summary>
        ''' <param name="StrConn">Connectionstring ou nome da conexao no webconfig (ex: "STRTAREFA", "ProviderName:MySQL.Data.MySQLClient;Server:127.0.0.1;Database:data;Uid:usuario;Pwd:senha;" ou "STRTAREFA;USER:usuario;PASSWORD:senha").</param>
        ''' <returns>Objeto connectionstring instanciado.</returns>
        ''' <remarks>Caso seja passada string, programador poderá fazer uso de complementos do tipo: "strGerador;user:estagiario;password:estag".
        ''' Isso corresponde a obter os dados da conexão do WebConfig com nome de strGerador e substituir nesta user e password.</remarks>
        Shared Function StrConnObj(ByVal StrConn As Object, ByVal ParamArray Params() As Object) As System.Configuration.ConnectionStringSettings
            If TypeOf (StrConn) Is System.Configuration.ConnectionStringSettings Then
                Return CType(StrConn, System.Configuration.ConnectionStringSettings)
            End If

            Dim Param As String = CType(StrConn, String)
            If Param.IndexOf(";") = -1 Then
                Return WebConn(Param)
            End If

            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Try
                MacroSubstSQL(Param, Nothing, ListaParametros)
            Catch
            End Try


            Dim Elem As New ElementosStr(Param, ";")
            Dim Conn As New System.Configuration.ConnectionStringSettings

            If Elem.Items("").Conteudo <> "" Then
                Dim ConnAnt As System.Configuration.ConnectionStringSettings = System.Configuration.ConfigurationManager.ConnectionStrings(Elem.Items("").Conteudo)
                Conn.ProviderName = ConnAnt.ProviderName
                Conn.ConnectionString = ConnAnt.ConnectionString
            End If

            If Elem.Exists("ProviderName") Then
                Conn.ProviderName = Elem.Items("ProviderName").Conteudo
                Elem.Items("ProviderName").Conteudo = Nothing
            End If

            If Compare(Conn.ProviderName, Oracle) Then
                If Elem.Exists("User") Then
                    Elem.Items("User").Nome = "User ID"
                End If
            ElseIf Compare(Conn.ProviderName, MySQL) Then
                If Elem.Exists("User") Then
                    Elem.Items("User").Nome = "Uid"
                End If
                If Elem.Exists("Password") Then
                    Elem.Items("Password").Nome = "Pwd"
                End If
            ElseIf Compare(Conn.ProviderName, MSAccess) Then
                If Elem.Exists("Data Source") Then
                    Dim Caminho As String = Elem.Items("Data Source").Conteudo
                    If Caminho.StartsWith("~/") Or Caminho.StartsWith("~\") Then
                        Caminho = HttpContext.Current.Server.MapPath(Caminho)
                    End If
                    Elem.Items("Data Source").Conteudo = Caminho
                End If
            End If

            Dim ElemNovo As New ElementosStr(Conn.ConnectionString, ";", "=")
            ElemNovo.AddStr(Elem.ToStyleStr(";", "="))
            Param = ElemNovo.ToStyleStr
            Conn.ConnectionString = Param

            Return Conn
        End Function

        ''' <summary>
        ''' Obtém string formatada considerando SELECT e DELETE ref à tabela a ser consultada, filtro e expressão de ordenação.
        ''' </summary>
        ''' <param name="Expressao">Expressão a ser consultada.</param>
        ''' <param name="TabelaouSQL">Tabela ou SQL a ser executado para obtenção da expressão.</param>
        ''' <param name="Filtro">Filtro no formato WHERE SQL.</param>
        ''' <returns>Retorna string contendo SELECT EXPRESSÃO FROM TABELAOUSQL WHERE FILTRO.</returns>
        ''' <remarks></remarks>
        Shared Function ExprSQL(ByVal TabelaouSQL As String, Optional ByVal Expressao As String = "*", Optional ByVal Filtro As String = "", Optional ByVal Ordem As String = "", Optional ByVal Tipo As ExprSQLTipo = ExprSQLTipo.Sel) As String
            Dim Pos As Integer = InStr(TabelaouSQL, " FROM ", CompareMethod.Text)
            Dim Sql As String = ""
            If Pos = 0 Then
                Sql = "FROM " & TabelaouSQL
            Else
                Sql = Mid(TabelaouSQL, Pos + 1)
            End If
            Sql &= IIf(Filtro <> "", " WHERE " & Filtro, "")
            Sql &= IIf(Ordem <> "", " ORDER BY " & Ordem, "")
            Return Microsoft.VisualBasic.Switch(Tipo = ExprSQLTipo.Del, "DELETE " & Sql, Tipo = ExprSQLTipo.Sel, "SELECT " & Expressao & " " & Sql, True, Nothing)
        End Function

        ''' <summary>
        ''' Carrega estrutura de um dataset baseado em SQL em ORACLE, MySQL ou MSAccess.
        ''' </summary>
        ''' <param name="SQL">Select para obtenção da estrutura.</param>
        ''' <param name="StrConn">Identificador da connexão ou string da configuração no web.config.</param>
        ''' <returns>Retorna um dataset contendo somente a estrutura.</returns>
        ''' <remarks></remarks>
        Shared Function DSCarregaEstrut(ByVal SQL As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object) As DataSet
            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

            Dim ds As DataSet = New DataSet
            If Compare(ConnW.ProviderName, MySQL) Then
                ' mysql
                Dim c As New CriadorDeObjetos("MySql.Data.dll")

                Dim Conexao As Object = c.Criar("MySqlConnection", ConnW.ConnectionString)
                Dim comm As Object = c.Criar("MySqlCommand")
                comm = DSCriaComandoMySQL(SQL, Conexao, ListaParametros)

                Dim Adapt As Object = c.Criar("MySqlDataAdapter", comm)
                Adapt.FillSchema(ds, SchemaType.Mapped)
                comm.Connection.Close()
            ElseIf Compare(ConnW.ProviderName, MSAccess) Then
                ' msaccess
                Dim Conexao As OleDbConnection = New OleDbConnection(ConnW.ConnectionString)
                Dim comm As OleDbCommand = DSCriaComandoAccess(SQL, Conexao, ListaParametros)
                Dim Adapt As OleDbDataAdapter = New OleDbDataAdapter(comm)
                Adapt.FillSchema(ds, SchemaType.Mapped)
                comm.Connection.Close()
            ElseIf Compare(ConnW.ProviderName, Oracle) Then
                ' oracle
                Dim c As New CriadorDeObjetos("System.Data.OracleClient.dll")
                Dim conexao As Object = c.Criar("OracleConnection", ConnW.ConnectionString)
                Dim comm As Object = c.Criar("OracleCommand")
                comm = DSCriaComandoOracle(SQL, conexao, ListaParametros)
                Dim Adapt As Object = c.Criar("OracleDataAdapter", comm)
                Adapt.FillSchema(ds, SchemaType.Mapped)
                comm.Connection.Close()
            End If
            Return ds
        End Function

        ''' <summary>
        ''' Carrega dados de um select considerando uma quantidade máxima de registros.
        ''' </summary>
        ''' <param name="Top">Número de registros a serem carregados a partir do topo.</param>
        ''' <param name="SQL">Expressão sql.</param>
        ''' <param name="StrConn">Identificador da conexão ou string de configuração no web.config.</param>
        ''' <param name="Params">Sequência e parâmetros do tipo ":campo1", conteudo1, ":campo2", conteudo2.</param>
        ''' <returns>Retorna um dataset contendo os dados dos registros selecionados.</returns>
        ''' <remarks>Primeiro parâmetro da lista também poderá conter um paramarray contendo todos os parâmetros.</remarks>
        Shared Function DSCarregaTop(ByVal Top As Integer, ByVal SQL As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object) As DataSet
            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)
            If Top > 0 Then
                If String.Compare(ConnW.ProviderName, MySQL, True) = 0 Then
                    SQL = "SELECT * FROM (" & SQL & ") as b LIMIT " & Top
                ElseIf String.Compare(ConnW.ProviderName, MSAccess, True) = 0 Then
                    SQL = "SELECT TOP " & Top & " * FROM (" & SQL & ")"
                ElseIf String.Compare(ConnW.ProviderName, Oracle, True) = 0 Then
                    SQL = "SELECT * FROM (" & SQL & ") WHERE ROWNUM < " & Top
                End If
            End If
            Return DSCarrega(SQL, ConnW, ListaParametros)
        End Function

        ''' <summary>
        ''' Carrega dados de um select.
        ''' </summary>
        ''' <param name="SQL">Expressão sql.</param>
        ''' <param name="StrConn">Identificador da string de conexão configurada no web.config.</param>
        ''' <param name="Params">Sequência e parâmetros do tipo ":campo1", conteudo1, ":campo2", conteudo2.</param>
        ''' <returns>Retorna um dataset contendo os dados dos registros selecionados.</returns>
        ''' <remarks>Primeiro parâmetro da lista também poderá conter um paramarray contendo todos os parâmetros.</remarks>
        Shared Function DSCarrega(ByVal SQL As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object) As DataSet
            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)
            Dim ds As DataSet = New DataSet

            If String.Compare(ConnW.ProviderName, MySQL, True) = 0 Then
                ' mysql
                Dim c As New CriadorDeObjetos("MySql.Data.dll")

                Dim Conexao As Object = c.Criar("MySqlConnection", ConnW.ConnectionString)
                Dim pref As String = New String("")

                ' inclui variáveis de sessão caso estejam definidas

#If Not _MyType = "WindowsForms" Then
                Dim Ctx As HttpContext = HttpContext.Current

                'Dim Pref As String = "SET @CONN_IP = " & SqlExpr(Ctx.Request.UserHostAddress) & ";"
                pref = "SET @CONN_IP = " & SqlExpr(Ctx.Request.UserHostAddress) & ";"

                pref &= "SET @CONN_MACHINE = " & SqlExpr(Ctx.Request.UserHostName) & ";"
                If Ctx.Session("CONN_USER") <> "" Then
                    pref &= "SET @CONN_USER = " & SqlExpr(Ctx.Session("CONN_USER")) & ";"
                End If
#End If

                Dim Comm As Object = c.Criar("MySqlCommand")
                Comm = DSCriaComandoMySQL(pref & SQL, Conexao, ListaParametros)

                Dim Adapt As Object = c.Criar("MySqlDataAdapter", Comm)
                Adapt.Fill(ds)
                Comm.Connection.Close()




            ElseIf String.Compare(ConnW.ProviderName, SQLServer, True) = 0 Then
                'sqlserver
                Dim Conexao As SqlConnection = New SqlConnection(ConnW.ConnectionString)
                Dim Comm As SqlCommand = DSCriaComandoSQLServer(SQL, Conexao, ListaParametros)
                Dim Adapt As SqlDataAdapter = New SqlDataAdapter(Comm)
                Adapt.Fill(ds)
                Comm.Connection.Close()



            ElseIf String.Compare(ConnW.ProviderName, MSAccess, True) = 0 Then
                ' msaccess
                Dim Conexao As OleDbConnection = New OleDbConnection(ConnW.ConnectionString)
                Dim Comm As OleDbCommand = DSCriaComandoAccess(SQL, Conexao, ListaParametros)
                Dim Adapt As OleDbDataAdapter = New OleDbDataAdapter(Comm)
                Adapt.Fill(ds)
                Comm.Connection.Close()
            ElseIf String.Compare(ConnW.ProviderName, Oracle, True) = 0 Then
                ' oracle
                Dim c As New CriadorDeObjetos("System.Data.OracleClient.dll")
                Dim conexao As Object = c.Criar("OracleConnection", ConnW.ConnectionString)
                Dim pref As New StringBuilder()

#If _MyType = "Web" Then
                Dim ctx As HttpContext = HttpContext.Current
#End If

                Dim comm As Object = c.Criar("OracleCommand")
                comm = DSCriaComandoOracle(pref.ToString() + SQL, conexao, ListaParametros)
                Dim Adapt As Object = c.Criar("OracleDataAdapter", comm)
                Adapt.Fill(ds)
                comm.Connection.Close()
            End If
            Return ds
        End Function

        ''' <summary>
        ''' Executa um comando de gravação conforme sql e parâmetros.
        ''' </summary>
        ''' <param name="SQL">Expressão sql.</param>
        ''' <param name="StrConn">Identificador da string de conexão configurada no web.config.</param>
        ''' <param name="Params">Sequência e parâmetros do tipo ":campo1", conteudo1, ":campo2", conteudo2.</param>
        ''' <remarks>Primeiro parâmetro da lista também poderá conter um paramarray contendo todos os parâmetros.</remarks>
        Shared Sub DSGrava(ByVal SQL As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object)
            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

            If String.Compare(ConnW.ProviderName, MySQL, True) = 0 Then
                ' mysql
                Dim c As New CriadorDeObjetos("MySql.Data.dll")

                Dim Conexao As Object = c.Criar("MySqlConnection", ConnW.ConnectionString)
                Dim Pref As String = New String("")

                ' inclui variáveis de sessão caso estejam definidas
#If _MyType = "Web" Then
                Dim Ctx As HttpContext = HttpContext.Current
                Pref = "SET @CONN_IP = " & SqlExpr(Ctx.Request.UserHostAddress) & ";"
                Pref &= "SET @CONN_MACHINE = " & SqlExpr(Ctx.Request.UserHostName) & ";"
                If Ctx.Session("CONN_USER") <> "" Then
                    Pref &= "SET @CONN_USER = " & SqlExpr(Ctx.Session("CONN_USER")) & ";"
                End If
#End If


                Dim Comm As Object = DSCriaComandoMySQL(Pref & SQL, Conexao, ListaParametros)
                Conexao.Open()
                Comm.ExecuteNonQuery()
                Conexao.Close()

            ElseIf String.Compare(ConnW.ProviderName, MSAccess, True) = 0 Then
                ' msaccess
                Dim Conexao As OleDbConnection = New OleDbConnection(ConnW.ConnectionString)
                Dim Comm As OleDbCommand = DSCriaComandoAccess(SQL, Conexao, ListaParametros)
                Try
                    Conexao.Open()
                    Comm.ExecuteNonQuery()
                    Conexao.Close()
                Catch ex As Exception
                    Dim txterr As String = ""
                    For Each param As OleDbParameter In Comm.Parameters
                        txterr &= param.ParameterName & " = " & NZ(param.Value, "") & " (" & param.DbType & ")" & vbCrLf
                    Next
                    txterr &= "(" & ex.Message & ")"
                    Throw New Exception("Erro ao tentar gravar " & Comm.CommandText & vbCrLf & txterr)
                End Try

            ElseIf String.Compare(ConnW.ProviderName, Oracle, True) = 0 Then
                ' oracle
                Dim c As New CriadorDeObjetos("System.Data.OracleClient.dll")
                Dim Conexao As Object = c.Criar("OracleConnection", ConnW.ConnectionString)  'OracleConnection = New OracleConnection(ConnW.ConnectionString)
                Dim pref As New StringBuilder("BEGIN ")

#If _MyType = "Web" Then
                Dim ctx As HttpContext = HttpContext.Current

                Dim conn_machine As String = SqlExpr(ctx.Request.UserHostName)
                Dim conn_ip As String = SqlExpr(ctx.Request.UserHostAddress)
                Dim conn_user As String = "''" 'Por enquanto vazio para garantir preenchimento

                If ctx.Session("CONN_USER") <> "" Then
                    conn_user = SqlExpr(ctx.Session("CONN_USER"))
                End If

                pref.Append("dbms_application_info.set_module(module_name => " & NZV(conn_machine, "''") & ", action_name => " & conn_user & ");")
                pref.Append("dbms_application_info.set_client_info(client_info => " & NZV(conn_ip, "''") & ");")
#End If

                pref.Append(IIf(SQL.EndsWith(";"), SQL, SQL & ";") & " END;")

                Dim Comm As Object = c.Criar("OracleCommand")
                Comm = DSCriaComandoOracle(pref.ToString, Conexao, ListaParametros)
                Conexao.Open()
                Comm.ExecuteNonQuery()
                Conexao.Close()
            End If
        End Sub

        ''' <summary>
        ''' Executa comandos de gravação conforme lista de sql e parâmetros.
        ''' </summary>
        ''' <param name="listaSql">Lista de instruções sql.</param>
        ''' <param name="StrConn">Identificador da string de conexão configurada no web.config.</param>
        ''' <param name="Params">Sequência e parâmetros do tipo ":campo1", conteudo1, ":campo2", conteudo2.</param>
        ''' <remarks></remarks>
        Shared Sub DSGrava(ByVal listaSql As IList, ByVal StrConn As Object, ByVal ParamArray Params() As Object)
            Dim comandos As String = Join(listaSql.OfType(Of String).ToArray, ";").Replace(";;", ";")

            DSGrava(comandos, StrConn, Params)
        End Sub

        ''' <summary>
        ''' Obtém ou define parâmetro padronizado em tabela.
        ''' </summary>
        ''' <param name="StrConn">Indicador de conexão (connection string).</param>
        ''' <param name="Chave">Chave para leitura ou gravação.</param>
        ''' <param name="Tabela">Tabela de parâmetros. Por default é SYS_CONFIG_GLOBAL.</param>
        ''' <param name="CampoChave">Campo da tabela que armazena a chave.</param>
        ''' <param name="CampoConteudo">Campo da tabela que armazena o conteúdo.</param>
        ''' <value>Conteúdo do parâmetro para gravação.</value>
        ''' <returns>Conteúdo do parâmetro lido.</returns>
        ''' <remarks></remarks>
        Shared Property DSConfig(ByVal StrConn As Object, ByVal Chave As String, Optional ByVal Tabela As Object = "SYS_CONFIG_GLOBAL", Optional ByVal CampoChave As String = "PARAM", Optional ByVal CampoConteudo As String = "CONFIG", Optional ByVal Params As Object = Nothing, Optional ByVal Params2 As Object = Nothing) As Object
            Get
                Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params, Params2)
                Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)
                Return NZ(DSValor(CampoConteudo, Tabela, ConnW, CampoChave & "=:PARAM", ":PARAM", Chave), "")
            End Get
            Set(ByVal value As Object)
                Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params, Params2)
                Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)
                If DSValor("COUNT(*)", Tabela, ConnW, CampoChave & "=:PARAM", ":PARAM", Chave) > 0 Then
                    DSGrava("UPDATE " & Tabela & " SET " & CampoConteudo & "=:CONTEUDO WHERE " & CampoChave & "=:PARAM", ConnW, ":CONTEUDO", NZ(value, ""), ":PARAM", Chave)
                Else
                    DSGrava("INSERT INTO " & Tabela & " (" & CampoChave & ", " & CampoConteudo & ") values(:CHAVE, :CONTEUDO)", ConnW, ":CHAVE", Chave, ":CONTEUDO", NZ(value, ""))
                End If
            End Set
        End Property

        ''' <summary>
        ''' Obtém ou define parâmetro de usuário específico em tabela de configuração.
        ''' </summary>
        ''' <param name="StrConn">Indicador de conexão (connection string).</param>
        ''' <param name="Usuario">Nome do usuário para o qual será atribuído ou obtido o parâmetro.</param>
        ''' <param name="Chave">Chave para leitura ou gravação.</param>
        ''' <param name="Tabela">Tabela de parâmetros. Por default é SYS_CONFIG_USUARIO.</param>
        ''' <param name="CampoChave">Campo da tabela que armazena a chave.</param>
        ''' <param name="CampoConteudo">Campo da tabela que armazena o conteúdo.</param>
        ''' <value>Conteúdo do parâmetro para gravação.</value>
        ''' <returns>Conteúdo do parâmetro lido.</returns>
        ''' <remarks></remarks>
        Shared Property DSConfigUsuario(ByVal StrConn As Object, ByVal Usuario As String, ByVal Chave As String, Optional ByVal Tabela As String = "SYS_CONFIG_USUARIO", Optional ByVal CampoUsuario As String = "USUARIO", Optional ByVal CampoChave As String = "PARAM", Optional ByVal CampoConteudo As String = "CONFIG", Optional ByVal Params As Object = Nothing, Optional ByVal Params2 As Object = Nothing) As Object
            Get
                Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params, Params2)
                Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)
                Return NZ(DSValor(CampoConteudo, Tabela, ConnW, CampoChave & "=:PARAM AND " & CampoUsuario & "=:USUARIO", ":PARAM", Chave, ":USUARIO", Usuario), "")
            End Get
            Set(ByVal value As Object)
                Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params, Params2)
                Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)
                If DSValor("COUNT(*)", Tabela, ConnW, CampoChave & "=:PARAM", ":PARAM", Chave) > 0 Then
                    DSGrava("UPDATE " & Tabela & " SET " & CampoConteudo & "=:CONTEUDO WHERE " & CampoChave & "=:PARAM AND " & CampoUsuario & "=:USUARIO", ConnW, ":CONTEUDO", NZ(value, ""), ":PARAM", Chave, ":USUARIO", Usuario)
                Else
                    If Usuario <> "" Then
                        DSGrava("INSERT INTO " & Tabela & " (" & CampoUsuario & ", " & CampoChave & ", " & CampoConteudo & ") values(:USUARIO, :CHAVE, :CONTEUDO)", ConnW, ":USUARIO", Usuario, ":CHAVE", Chave, ":CONTEUDO", NZ(value, ""))
                    End If
                End If
            End Set
        End Property

        ''' <summary>
        ''' Obtem um valor em uma tabela do tipo max(seq) ou count(seq), por exemplo.
        ''' </summary>
        ''' <param name="Expressao">Expressão como min(campo), max(campo) ou count(campo), por exemplo.</param>
        ''' <param name="TabelaouSQL">Tabela ou consulta a ser pesquisada.</param>
        ''' <param name="StrConn">Identificador da string de conexão configurada no web.config ou a própria conexão.</param>
        ''' <param name="Condicao">Filtro a ser aplicado na cláusula where da expressão select.</param>
        ''' <param name="Params">Sequência e parâmetros do tipo ":campo1", conteudo1, ":campo2", conteudo2.</param>
        ''' <returns>Retorna o valor de uma expressão na primeira linha do dataset resultante.</returns>
        ''' <remarks>Primeiro parâmetro da lista também poderá conter um paramarray contendo todos os parâmetros.</remarks>
        Shared Function DSValor(ByVal Expressao As String, ByVal TabelaouSQL As String, ByVal StrConn As Object, ByVal Condicao As String, ByVal ParamArray Params() As Object) As Object
            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

            Dim SQL As String = ExprSQL(TabelaouSQL, Expressao & " AS VRESULT", Condicao)
            Dim DS As DataSet = DSCarrega(SQL, ConnW, ListaParametros)
            If DS.Tables(0).Rows.Count = 1 Then

                ' RETORNA VALOR
                Return DS.Tables(0).Rows(0)("VRESULT")
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Retorna próxima sequência de um campo numérico ou referencial.
        ''' </summary>
        ''' <param name="Campo">Campo de auto-sequenciação.</param>
        ''' <param name="TabelaouSQL">Nome da tabela.</param>
        ''' <param name="StrConn">Identificador da string de conexão configurada no web.config ou a própria conexão.</param>
        ''' <param name="Condicao">Filtro a ser aplicado na cláusula where da expressão select.</param>
        ''' <param name="Params">Sequência e parâmetros do tipo ":campo1", conteudo1, ":campo2", conteudo2.</param>
        ''' <returns>Retorna o próxima sequência.</returns>
        ''' <remarks>Primeiro parâmetro da lista também poderá conter um paramarray contendo todos os parâmetros.</remarks>
        Shared Function DSProxSeq(ByVal Campo As String, ByVal TabelaouSQL As String, ByVal StrConn As Object, ByVal Condicao As String, ByVal ParamArray Params() As Object) As Object
            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

            Dim SQL As String = ExprSQL(TabelaouSQL, " MAX(" & Campo & ") AS VRESULT", Condicao)
            Dim DS As DataSet = DSCarrega(SQL, ConnW, ListaParametros)
            If DS.Tables(0).Rows.Count = 1 Then

                ' RETORNA VALOR
                Return NZ(DS.Tables(0).Rows(0)("VRESULT"), 0) + 1
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Constrói arraylist com base em expressão selecionada na tabela.
        ''' </summary>
        ''' <param name="Expressao">Campo ou expressão</param>
        ''' <param name="TabelaouSQL">Tabela ou SQL onde será a consulta.</param>
        ''' <param name="StrConn">Identificador da string de conexão configurada no web.config ou a própria conexão.</param>
        ''' <param name="Condicao">Filtro a ser aplicado na cláusula where da expressão select.</param>
        ''' <param name="Ordem">Campos para ordenação separados por vírgula.</param>
        ''' <param name="Params">Sequência e parâmetros do tipo ":campo1", conteudo1, ":campo2", conteudo2.</param>
        ''' <returns>Retorna um arraylist contendo como item o resultado da expressão por registro.</returns>
        ''' <remarks>Primeiro parâmetro da lista também poderá conter um paramarray contendo todos os parâmetros.</remarks>
        Shared Function DSConcat(ByVal Expressao As String, ByVal TabelaouSQL As String, ByVal StrConn As Object, ByVal Condicao As String, ByVal Ordem As String, ByVal ParamArray Params() As Object) As ArrayList
            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

            Dim SQL As String = ExprSQL(TabelaouSQL, Expressao & " AS VRESULT", Condicao, Ordem)
            Dim DS As DataSet = DSCarrega(SQL, ConnW, ListaParametros)
            Return ItemsToArrayList(DS.Tables(0).Rows, "VRESULT")
        End Function

        ''' <summary>
        ''' Cria comando para MySQL.
        ''' </summary>
        ''' <param name="SQL">SQL a ser configurado com parâmetros.</param>
        ''' <param name="Conexao">Conexão para acesso ao banco de dados.</param>
        ''' <param name="ListaParametros">Lista de parâmetros tipo ParamArray.</param>
        ''' <returns>Retorna comando a ser executado.</returns>
        ''' <remarks></remarks>
        Private Shared Function DSCriaComandoMySQL(ByRef SQL As String, ByRef Conexao As Object, ByRef ListaParametros As ArrayList) As Object
            Dim c As New CriadorDeObjetos("MySql.Data.dll")
            MacroSubstSQL(SQL, ListaParametros)
            Dim Comando As Object = c.Criar("MySqlCommand", SQL, Conexao)
            Dim Param As String
            Dim gexmatch As MatchCollection = Regex.Matches(SQL, "[:@\?](\w+)", RegexOptions.Multiline)
            For Each m As Match In gexmatch
                Param = m.Value
                Dim pos As Integer = ListaParametros.IndexOf(m.Value)
                If pos <> -1 Then
                    Mid(Comando.CommandText, m.Groups(1).Index, 1) = "?"
                    Mid(Param, 1, 1) = "?"
                    Dim P As Object
                    P = c.Criar("MySqlParameter", Param, ListaParametros(pos + 1))
                    Comando.Parameters.Add(P)
                End If
            Next
            Return Comando
        End Function

        ''' <summary>
        ''' Constrói comando para Microsoft Access.
        ''' </summary>
        ''' <param name="SQL">SQL a ser executado com parâmetros.</param>
        ''' <param name="Conexao">Conexão com a base de dados.</param>
        ''' <param name="ListaParametros">Lista de parâmetros tipo ParamArray.</param>
        ''' <returns>Retorna comando a ser executado.</returns>
        ''' <remarks></remarks>
        Private Shared Function DSCriaComandoAccess(ByRef SQL As String, ByRef Conexao As OleDbConnection, ByRef ListaParametros As ArrayList) As OleDbCommand
            MacroSubstSQL(SQL, ListaParametros)
            SQL = SQL.Replace(" || ", " & ")
            Dim Comando As OleDbCommand = New OleDbCommand(SQL, Conexao)
            Dim Param As String
            Dim gexmatch As MatchCollection = Regex.Matches(SQL, "[:@\?](\w+)", RegexOptions.Multiline)
            For Each m As Match In gexmatch
                Param = m.Groups(1).Value
                Dim pos As Integer = ListaParametros.IndexOf(m.Value)
                If pos <> -1 Then
                    Mid(Comando.CommandText, m.Groups(1).Index, 1) = "@"
                    Param = "@" & Param
                    Dim P As OleDbParameter
                    If IsNothing(ListaParametros(pos)) Then
                        P = New OleDbParameter(Param, ListaParametros(pos + 1))
                    ElseIf ListaParametros(pos + 1).GetType.ToString = "System.DateTime" Then
                        P = New OleDbParameter(Param, OleDb.OleDbType.Date)
                        P.Value = ListaParametros(pos + 1)
                    Else
                        P = New OleDbParameter(Param, ListaParametros(pos + 1))
                    End If
                    Comando.Parameters.Add(P)
                End If
            Next
            Return Comando
        End Function

        ''' <summary>
        ''' Cria comando para Oracle.
        ''' </summary>
        ''' <param name="SQL">Comando SQL a ser executado com parâmetros.</param>
        ''' <param name="Conexao">Conexão com a base de dados.</param>
        ''' <param name="ListaParametros">Lista de parâmetros no formato ParamArray.</param>
        ''' <returns>Retorna comando a ser executado no adapter.</returns>
        ''' <remarks></remarks>
        Private Shared Function DSCriaComandoOracle(ByRef SQL As String, ByRef Conexao As Object, ByRef ListaParametros As ArrayList) As Object
            Dim c As New CriadorDeObjetos("System.Data.OracleClient.dll")
            MacroSubstSQL(SQL, ListaParametros)
            Dim Comando As Object = c.Criar("OracleCommand", SQL, Conexao)
            Dim Param As String
            Dim gexmatch As MatchCollection = Regex.Matches(SQL, "[:@\?](\w+)", RegexOptions.Multiline)
            For Each m As Match In gexmatch
                Param = m.Value
                Dim pos As Integer = ListaParametros.IndexOf(m.Value)
                If pos <> -1 Then
                    Mid(Comando.CommandText, m.Groups(1).Index, 1) = ":"
                    Mid(Param, 1, 1) = ":"
                    Dim P As Object
                    P = c.Criar("OracleParameter", Param, ListaParametros(pos + 1))
                    Comando.Parameters.Add(P)
                End If
            Next
            Return Comando
        End Function


        ''' <summary>
        ''' Cria comando para SQL Server.
        ''' </summary>
        ''' <param name="SQL">Comando SQL a ser executado com parâmetros.</param>
        ''' <param name="Conexao">Conexão com a base de dados.</param>
        ''' <param name="ListaParametros">Lista de parâmetros no formato ParamArray.</param>
        ''' <returns>Retorna comando a ser executado no adapter.</returns>
        ''' <remarks></remarks>
        Private Shared Function DSCriaComandoSQLServer(ByRef SQL As String, ByRef Conexao As SqlConnection, ByRef ListaParametros As ArrayList) As SqlCommand
            MacroSubstSQL(SQL, ListaParametros)
            Dim Comando As SqlCommand = New SqlCommand(SQL, Conexao)
            Dim Param As String
            Dim gexmatch As MatchCollection = Regex.Matches(SQL, "[:@\?](\w+)", RegexOptions.Multiline)

            For Each m As Match In gexmatch
                Param = m.Value
                Dim pos As Integer = ListaParametros.IndexOf(m.Value)
                If pos <> -1 Then
                    Mid(Comando.CommandText, m.Groups(1).Index, 1) = "@"
                    Mid(Param, 1, 1) = "@"
                    Dim P As SqlParameter
                    P = New SqlParameter(Param, ListaParametros(pos + 1))
                    Comando.Parameters.Add(P)
                End If
            Next
            Return Comando
        End Function

        ''' <summary>
        ''' Retorna tipo da base de dados.
        ''' </summary>
        ''' <param name="StrConn">Identificador do connection.</param>
        ''' <param name="Params">Parâmetros para serem utilizados por macrosubst.</param>
        ''' <returns>Retorna enum TipoBaseSQL.MSAccess, MySQL ou Oracle.</returns>
        ''' <remarks></remarks>
        Shared Function DSTipoBaseSQL(ByVal StrConn As Object, ByVal ParamArray Params() As Object) As TipoBaseSQL
            Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)
            If Compare(ConnW.ProviderName, MySQL) Then
                Return TipoBaseSQL.MySQL
            ElseIf Compare(ConnW.ProviderName, MSAccess) Then
                Return TipoBaseSQL.MSAccess
            ElseIf Compare(ConnW.ProviderName, Oracle) Then
                Return TipoBaseSQL.Oracle
            ElseIf Compare(ConnW.ProviderName, SQLServer) Then
                Return TipoBaseSQL.SQLServer
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Obtém um DataColumn com base em um DataTable.
        ''' </summary>
        ''' <param name="DSTab">DataTable que será utilizado como base para se obter as colunas.</param>
        ''' <param name="CamposTxt">Campos desejados do DataTable separados por ;.</param>
        ''' <returns>Retorna um DataColumn contendo os campos passados em CamposTxt que existam dentro de DSTab.</returns>
        ''' <remarks></remarks>
        Shared Function DSDataColumns(ByVal DSTab As DataTable, ByVal CamposTxt As String) As DataColumn()
            Dim Campos As Array = Split(CamposTxt, ";")
            Dim Cols() As DataColumn = New DataColumn(Campos.Length) {}
            For z As Integer = 0 To Campos.Length - 1
                Cols(z) = DSTab.Columns(Campos(z))
            Next
            Return Cols
        End Function

        ''' <summary>
        ''' Retorna um dataset filtrado
        ''' </summary>
        ''' <param name="DSCompleto">Dataset que se deseja filtrar.</param>
        ''' <param name="Filtro">String correspondente ao filtro.</param>
        ''' <returns>DataSet com o conjunto de registro correspondentes ao filtro passado.</returns>
        ''' <remarks></remarks>
        Shared Function DSFiltra(ByVal DSCompleto As DataSet, ByVal Filtro As String) As DataSet
            Return DSFiltra(DSCompleto, Filtro, "")
        End Function

        ''' <summary>
        ''' Retorna um dataset filtrado
        ''' </summary>
        ''' <param name="DSCompleto">Dataset que se deseja filtrar.</param>
        ''' <param name="Filtro">String correspondente ao filtro.</param>
        ''' <param name="Classifica">Coluna pela qual o DataSet será ordernado.</param>
        ''' <returns>DataSet com o conjunto de registro correspondentes ao filtro passado.</returns>
        ''' <remarks></remarks>
        Shared Function DSFiltra(ByVal DSCompleto As DataSet, ByVal Filtro As String, ByVal Classifica As String) As DataSet

            Dim DrFiltrada() As DataRow = DSCompleto.Tables(0).Select(Filtro, Classifica)
            Dim DsFiltrado As DataSet
            DsFiltrado = DSCompleto.Clone

            For r As Integer = 0 To DrFiltrada.Length - 1
                DsFiltrado.Tables(0).ImportRow(DrFiltrada(r))
            Next

            Return DsFiltrado

        End Function


        ''' <summary>
        ''' Conversão de estrutura recursiva em tabela para hierarquia em texto tipo arraylist.
        ''' </summary>
        ''' <param name="DS">Dataset contendo recursividade.</param>
        ''' <param name="Chave">Campo chave da tabela.</param>
        ''' <param name="CampoVinc">Campo vinculado que caracteriza recursividade.</param>
        ''' <param name="MascNode">Campo que será apresentado como texto do node.</param>
        ''' <param name="Filtro">Filtro para pesquisa.</param>
        ''' <returns>Arraylist contendo hierarquia.</returns>
        ''' <remarks></remarks>
        Shared Function TabHierarqParaMenu(ByVal DS As DataSet, ByVal Chave As String, ByVal CampoVinc As String, ByVal MascNode As String, Optional ByVal Filtro As Object = Nothing) As ArrayList
            Dim Itens As ArrayList = New ArrayList
            Dim DV As DataView = New DataView(DS.Tables(0), CampoVinc & IIf(IsNothing(Filtro), " IS NULL", " = " & Filtro), "ORDEM, SEQ", DataViewRowState.CurrentRows)
            For Each Row As DataRowView In DV

                Dim Str As String = "'" & Row("DESCR") & "'"
                For Each Campo As String In Split("Objeto;URL;Prefix;Tip;Seq;Super_Seq", ";")
                    If NZ(Row(Campo), "") <> "" Then
                        Str &= " " & Campo & ":" & NZ(Row(Campo), "")
                    End If
                Next
                Itens.Add(Str)

                Dim SubItens As ArrayList = TabHierarqParaMenu(DS, Chave, CampoVinc, MascNode, Row("SEQ"))
                If SubItens.Count <> 0 Then
                    Itens.Add(SubItens)
                End If
            Next
            Return Itens
        End Function

        ''' <summary>
        ''' Concatena um conjunto de strings separadas ou não por um separador.
        ''' </summary>
        ''' <param name="Sep">Separador que será colocado entre as strings que serão concatenadas.</param>
        ''' <param name="Segmentos">Conjunto de strings que serão concatenadas.</param>
        ''' <returns>Retorna uma string única contendo todas as Strings presentes em Segmentos concatenadas e separadas por Sep.</returns>
        ''' <remarks></remarks>
        Public Shared Function StrExpr(ByVal Sep As String, ByVal ParamArray Segmentos() As Object) As String
            Dim Ret As String = ""
            For Each Item As String In Segmentos
                If Item <> "" Then
                    If Ret <> "" AndAlso Not (Ret.EndsWith(Sep) Or Item.StartsWith(Sep)) Then
                        Ret &= Sep
                    End If
                End If
            Next
            Do While Ret.StartsWith(Sep)
                Ret = Ret.Substring(Len(Sep))
            Loop
            Do While Ret.EndsWith(Sep)
                Ret = Ret.Substring(0, Len(Ret) - Len(Sep) - 1)
            Loop
            Return Ret
        End Function

        ''' <summary>
        ''' Busca na string formato de emailstr apenas o trecho com o email.
        ''' </summary>
        ''' <param name="Email">Email a ser analisado.</param>
        ''' <returns>Trecho contendo somente o email.</returns>
        ''' <remarks></remarks>
        Public Shared Function SoEmailStr(ByVal Email As String) As String
            Return RegexGroup(Email, "(^|[ \t\[\<\>\""]*)([\w-.]+@[\w-]+(\.[\w-]+)+)(($|[ \t\<\>\""]*))", 2).Value
        End Function

        ''' <summary>
        ''' Formata email conforme exigência dos servidores smtp, baseando-se no formato 'nome' [email@dominio.com.br].
        ''' </summary>
        ''' <param name="Email">Email a ser formatado.</param>
        ''' <returns>Retorna email no formato 'nome' &lt;email@dominio.com.br&gt;</returns>
        ''' <remarks></remarks>
        Public Shared Function EmailStr(ByVal Email As String) As String
            Email = Trim(Email)
            Email = Email.Replace("[", "<").Replace("]", ">").Replace(Chr(160), " ")
            If Email.StartsWith("'") Then
                Email = Regex.Replace(Email, "'(.*)'", """$1""")
            End If

            Email = Email.Replace("'", "`")

            If Email.IndexOf("<") = -1 Then
                Dim SoEmail As String = SoEmailStr(Email)
                If SoEmail <> "" Then
                    Email = ReplRepl(Email, SoEmail, "")
                End If
                Email = ReplRepl(Email, Chr(9), "")
                Email = Trim(ReplRepl(Email, "  ", " "))
                If Email <> "" Then
                    Email = SqlExpr(Email, """")
                End If
                Email = Email & " <" & SoEmail & ">"
            End If
            Return Email
        End Function

        ''' <summary>
        ''' Trata string transformando-a em arraylist de emails no formato 'nome'(email@dominio.com.br).
        ''' </summary>
        ''' <param name="Email">Email a ser tratado contendo lista separada por ponto e vírgula.</param>
        ''' <returns>Retorna arraylist contendo emails.</returns>
        ''' <remarks></remarks>
        Public Shared Function TermosStrToLista(ByVal Email As Object) As ArrayList
            If TypeOf (Email) Is ArrayList Then
                Email = Join(CType(Email, ArrayList).ToArray, ";")
            End If
            Dim Lista As New ArrayList
            If NZ(Email, "") <> "" Then
                Dim Emails As Array = Split(Join(Split(Email, vbCrLf), ";"), ";")
                For Each Item As String In Emails
                    Item = Trim(Item)
                    If Item <> "" Then
                        Dim pref As String
                        If Item.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                            pref = "bcc:"
                            Item = Item.Substring(4)
                        Else
                            pref = ""
                        End If
                        If Item.StartsWith("conf.", StringComparison.OrdinalIgnoreCase) Then
                            Dim Result As ArrayList = TermosStrToLista(WebConf(Item.Substring(5)))
                            If pref <> "" Then
                                For Each ResultItem As String In Result
                                    Lista.Add(pref & ResultItem)
                                Next
                            Else
                                Lista.AddRange(Result)
                            End If
                        Else
                            Lista.Add(pref & Item)
                        End If
                    End If
                Next
            End If
            Return Lista
        End Function

        ''' <summary>
        ''' Transforma texto em conteúdo cripto reversível com base em chave e algorítmo próprio.
        ''' </summary>
        ''' <param name="Texto">Texto a ser criptografado.</param>
        ''' <param name="Chave">Chave opcional. Na falta desta, default da biblioteca será utilizada.</param>
        ''' <returns>Retorna texto criptografado que representa conteúdo inicial, podendo ser revertido mediante chave.</returns>
        ''' <remarks></remarks>
        Shared Function EncrypB(ByVal Texto As String, Optional ByVal Chave As String = EncrypBChavePadrao) As String
            If Chave = "" Then
                Chave = EncrypBChavePadrao
            End If
            Dim Result As New List(Of Byte)
            For z As Integer = 0 To Len(Texto) - 1
                Dim pos As Integer = z Mod Len(Chave)
                Result.Add((Asc(Texto.Substring(z, 1)) + Asc(Chave.Substring(pos, 1))) Mod 256)
            Next
            Return Convert.ToBase64String(Result.ToArray)
        End Function

        ''' <summary>
        ''' Reversão de conteúdo criptografado com base em chave.
        ''' </summary>
        ''' <param name="Texto">Texto criptografado para reversão.</param>
        ''' <param name="Chave">Chave opcional. Na falta desta, default da biblioteca será considerada.</param>
        ''' <returns>Retorna conteúdo original, caso chave esteja correta.</returns>
        ''' <remarks></remarks>
        Shared Function DecrypB(ByVal Texto As String, Optional ByVal Chave As String = EncrypBChavePadrao) As String
            If Chave = "" Then
                Chave = EncrypBChavePadrao
            End If
            If IsNothing(Texto) OrElse IsDBNull(Texto) Then
                Texto = ""
            End If
            Dim Result As String = ""
            Dim Bytes() As Byte = Convert.FromBase64String(Texto)
            For z As Integer = 0 To Bytes.Length - 1
                Dim pos As Integer = z Mod Len(Chave)
                Dim Num As Integer = Bytes(z) - Asc(Chave.Substring(pos, 1))
                If Num < 0 Then
                    Num += 256
                End If
                Result &= Chr(Num)
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Reverte texto convertido por URLJSEncode, que tem finalidade de ocultar códigos especiais não tratáveis por GET.
        ''' </summary>
        ''' <param name="Texto">Texto codificado a ser revertido em original.</param>
        ''' <returns>Retorna texto revertido.</returns>
        ''' <remarks></remarks>
        Public Shared Function URLJSDecode(ByVal Texto As String) As String
            Dim pos As Integer = InStr(Texto, "_")
            Do While pos <> 0
                Texto = Microsoft.VisualBasic.Left(Texto, pos - 1) & Uri.HexUnescape("%" & Mid(Texto, pos + 1, 2), 0) & Mid(Texto, pos + 3)
                pos = InStr(pos + 1, Texto, "_")
            Loop
            Return Texto
        End Function

        ''' <summary>
        ''' Converte texto em codificação capaz de ocultar caracteres especiais com objetivo de passar entre JAVASCRIPT e ASP.NET.
        ''' </summary>
        ''' <param name="Texto">Texto original a ser convertido.</param>
        ''' <returns>Retorna texto codificado.</returns>
        ''' <remarks></remarks>
        Public Shared Function URLJSEncode(ByVal Texto As String) As String
            Dim ret As String = ""
            For z As Integer = 1 To Len(Texto)
                ret &= "_" & Mid(Uri.HexEscape(Mid(Texto, z, 1)), 2)
            Next
            Return ret
        End Function

        ''' <summary>
        ''' Obtém formato extendido, armazenado em texto contínuo e prepara arraylist com as variáveis seguidas de conteúdo (:var, :conteudo, :var2, conteudo2...).
        ''' </summary>
        ''' <param name="Extend">Texto a ser analisado.</param>
        ''' <returns>Arraylist contendo variáveis no formato :var1, conteudo1, :var2, conteudo2...</returns>
        ''' <remarks></remarks>
        Shared Function ExtendToArrayList(ByVal Extend As String) As ArrayList
            Dim Lista As New ArrayList
            If Extend.StartsWith("<<") Then
                Dim Var As String = ""
                Dim Conteudo As String = ""
                For Each linha As String In Split(Extend, vbCrLf)
                    If linha.StartsWith("<<") And linha.EndsWith(">>") Then
                        If Var <> "" Then
                            Lista.Add(Var)
                            Lista.Add(Conteudo)
                            Conteudo = ""
                        End If
                        Var = StrStr(linha, 2, -2)
                    Else
                        Conteudo &= IIf(Conteudo <> "", vbCrLf, "") & linha
                    End If
                Next
                If Var <> "" Then
                    Lista.Add(Var)
                    Lista.Add(Conteudo)
                End If
            Else
                Dim Elems As New ElementosStr(Extend, "|", ":")
                For Each Elem As ElementoStr In Elems.Elementos
                    Lista.Add(Elem.Nome)
                    Lista.Add(Elem.Conteudo)
                Next
            End If
            Return Lista
        End Function

        ''' <summary>
        ''' Transforma objeto em array de bytes pronto para gravar em blob.
        ''' </summary>
        ''' <param name="obj">Objeto a ser convertido em array de bytes.</param>
        ''' <returns>Retorna array de bytes.</returns>
        ''' <remarks></remarks>
        Shared Function ObjectToByteArray(ByVal obj As Object) As Byte()
            Dim Bytes() As Byte = Nothing
            Try
                Dim fs As System.IO.MemoryStream = New System.IO.MemoryStream
                Dim formatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                formatter.Serialize(fs, obj)
                Bytes = fs.ToArray
            Catch
            End Try
            Return Bytes
        End Function

        ''' <summary>
        ''' Obtém e seta conteúdo criptografado de propriedade.
        ''' </summary>
        ''' <param name="Objeto">Objeto para o qual a propriedade será tratada.</param>
        ''' <param name="Propriedade">Nome da propriedade.</param>
        ''' <param name="Container">Container quando objeto estiver em outro continente.</param>
        ''' <param name="ChaveEncrypB">Chave de criptografia sendo opcional. Na falta desta, default da biblioteca será utilizada.</param>
        ''' <value>Valor para definição da propriedade.</value>
        ''' <returns>Valor obtido com base na propriedade pesquisada.</returns>
        ''' <remarks></remarks>
        Shared Property PropE(ByVal Objeto As Object, Optional ByVal Propriedade As String = "", Optional ByVal Container As Object = Nothing, Optional ByVal ChaveEncrypB As String = EncrypBChavePadrao) As Object
            Get
                Return DecrypB(Prop(Objeto, EncrypB(Propriedade, ChaveEncrypB).Replace("=", ""), Container), ChaveEncrypB)
            End Get
            Set(ByVal value As Object)
                Prop(Objeto, EncrypB(Propriedade, ChaveEncrypB).Replace("=", ""), Container) = EncrypB(value, ChaveEncrypB)
            End Set
        End Property

        ''' <summary>
        ''' Define ou obter propriedade de um objeto ou resultado de uma busca em um container.
        ''' </summary>
        ''' <param name="Objeto">Objeto ou string contendo o nome do objeto a ser procurado.</param>
        ''' <param name="Propriedade">Nome da propriedade a ser definida.</param>
        ''' <param name="Container">Caso esse parâmetro não seja informado, propriedade do objeto será definida. Caso esteja presente, deverá indicar a página ou colleção de controles onde o nome do objeto será procurado.</param>
        ''' <value>Conteúdo a ser atribuído ao objeto ou resultado de pesquisa.</value>
        ''' <returns>Conteúdo obtido a partir do objeto ou resultado da pesquisa.</returns>
        ''' <remarks>Todas as propriedades precisam de prévia programação. Ao informar novidade em 'get', informe também em 'set' e vice-versa.</remarks>
        Shared Property Prop(ByVal Objeto As Object, Optional ByVal Propriedade As String = "", Optional ByVal Container As Object = Nothing) As Object
            Get

                ' busca pelo elemento ou retorno de NOTHING
                If IsNothing(Objeto) Then
                    Return Nothing
                End If
                Dim tipo As String = Objeto.GetType.ToString
                If IsNothing(Container) Then
                    If Compare(tipo, "System.String") Then
                        Return Nothing
                    End If
                Else
                    Objeto = Form.FindControl(Container, Objeto)
                    tipo = Objeto.GetType.ToString
                End If

                ' nenhum elemento, retorna nothing
                If IsNothing(Objeto) Then
                    Return Nothing
                End If

                ' tratamento de propriedade
                Propriedade = LCase(Propriedade)

                Try
                    Select Case Propriedade
                        Case "", "checked", "text", "valoratual", "conteudo", "value"
                            Select Case tipo
                                Case "System.Web.UI.WebControls.CheckBox", "System.Windows.Forms.CheckBox"
                                    Return NZV(Objeto.Checked, Objeto.page.request.form(Objeto.uniqueid))
                                Case "ASP.uc_icfttextarea_ascx"
                                    Return Objeto.Attributes("text")
                                Case "ASP.uc_icftgridview_ascx", "ASP.uc_icftdetalhes_ascx"
                                    Return Objeto.ChaveSel
                                Case "System.Web.UI.WebControls.HiddenField", "System.Xml.XmlAttribute"
                                    Return Objeto.Value
                                Case "Icraft+ElementoStr"
                                    Return Objeto.Conteudo
                                Case "System.Web.UI.WebControls.TextBox"
                                    Return NZ(Objeto.text, Objeto.page.request.form(Objeto.uniqueid))
                                Case "System.Web.UI.WebControls.DropDownList"
                                    Return NZ(Objeto.TEXT, Objeto.page.request.form(Objeto.uniqueid))
                                Case "ASP.uc_icftgridview_icftgridview_ascx"
                                    Return Objeto.chavesel
                                Case Else
                                    Return Objeto.Text
                            End Select
                        Case "enabled"
                            Return Objeto.Enabled
                        Case "backcolor"
                            Return Objeto.BackColor
                        Case "visible"
                            Return Objeto.Visible
                        Case "tooltip"
                            Return Objeto.ToolTip
                        Case "cssclass"
                            Return Objeto.CssClass
                        Case "validationgroup"
                            Return Objeto.ValidationGroup
                        Case "forecolor"
                            Return Objeto.ForeColor
                        Case "readonly"
                            Return Objeto.ReadOnly
                        Case "imageurl"
                            Return Objeto.ImageUrl
                        Case "nome", "name"
                            Select Case tipo
                                Case "Icraft+ElementoStr"
                                    Return Objeto.Nome
                                Case Else
                                    Return Objeto.Name
                            End Select
                        Case "id"
                            If tipo = "Icraft+ElementoStr" Then
                                Return Objeto.Nome
                            ElseIf TypeOf Objeto Is Windows.Forms.Control Then
                                Return CType(Objeto, Windows.Forms.Control).Name
                            End If
                            Return Objeto.ID
                        Case "navigateurl"
                            Return Objeto.NavigateUrl
                        Case Else
                            Try
                                Return Objeto.GetType.GetProperty(Propriedade, Reflection.BindingFlags.Public + Reflection.BindingFlags.Instance + Reflection.BindingFlags.IgnoreCase).GetValue(Objeto, Nothing)
                            Catch
                                Return Objeto.Attributes(Propriedade)
                            End Try
                    End Select
                Catch
                    Try
                        Return Objeto.Attributes(Propriedade)
                    Catch
                        Return Convert.DBNull
                    End Try
                End Try
            End Get

            Set(ByVal value As Object)

                ' busca pelo elemento ou retorno de NOTHING
                If IsNothing(Objeto) Then
                    Throw New Exception("Tentativa de definição de um objeto inexistente em Prop.")
                End If
                Dim Tipo As String = Objeto.GetType.ToString
                If IsNothing(Container) Then
                    If Compare(Tipo, "System.String") Then
                        Throw New Exception("Tentativa de definição de um objeto inexistente em Prop.")
                    End If
                Else
                    Objeto = Form.FindControl(Container, Objeto)
                    Tipo = Objeto.GetType.ToString
                End If

                ' nenhum elemento, gera erro
                If IsNothing(Objeto) Then
                    Throw New Exception("Tentativa de definição de um objeto inexistente em Prop.")
                End If

                ' tratamento de propriedade
                Propriedade = LCase(Propriedade)

                Try
                    Select Case Propriedade
                        Case "", "checked", "text", "valoratual", "conteudo", "value"
                            Select Case Tipo
                                Case "System.Web.UI.WebControls.CheckBox", "System.Windows.Forms.CheckBox"
                                    value = NZ(value, "")
                                    If value = "" OrElse value = "off" Then
                                        value = System.Boolean.FalseString
                                    ElseIf value = "on" Then
                                        value = System.Boolean.TrueString
                                    End If
                                    Objeto.Checked = value
                                Case "ASP.uc_icfttextarea_ascx"
                                    Objeto.Attributes("text") = value
                                Case "ASP.uc_icftgridview_icftgridview_ascx"
                                    Objeto.ChaveSel = value
                                Case "ASP.uc_icftgridview_ascx", "ASP.uc_icftdetalhes_ascx"
                                    Objeto.ChaveSel = value
                                Case "System.Web.UI.WebControls.HiddenField", "System.Xml.XmlAttribute"
                                    Objeto.Value = value
                                Case "Icraft+ElementoStr"
                                    Objeto.Conteudo = value
                                Case "System.Web.UI.WebControls.DropDownList"
                                    If IsNothing(Objeto.Items.FindByValue(value)) Then
                                        Objeto.Items.Add(value)
                                    End If
                                    Objeto.Text = value
                                Case Else
                                    Objeto.Text = value
                            End Select
                        Case "enabled"
                            Objeto.Enabled = value
                        Case "backcolor"
                            Objeto.BackColor = value
                        Case "visible"
                            Objeto.Visible = value
                        Case "tooltip"
                            Objeto.ToolTip = value
                        Case "cssclass"
                            Objeto.CssClass = value
                        Case "validationgroup"
                            Objeto.ValidationGroup = value
                        Case "forecolor"
                            Objeto.ForeColor = value
                        Case "readonly"
                            Objeto.ReadOnly = value
                        Case "nome", "name"
                            Select Case Tipo
                                Case "Icraft+ElementoStr"
                                    Objeto.Nome = value
                                Case Else
                                    Objeto.Name = value
                            End Select
                        Case "id"
                            If Tipo = "Icraft+ElementoStr" Then
                                Objeto.Nome = value
                            ElseIf TypeOf Objeto Is Windows.Forms.Control Then
                                CType(Objeto, Windows.Forms.Control).Name = value
                            End If
                            Objeto.ID = value
                        Case "navigateurl"
                            Objeto.NavigateUrl = value
                        Case "imageurl"
                            Objeto.ImageUrl = value
                        Case Else
                            Objeto.Attributes(Propriedade) = value
                    End Select
                Catch
                    Objeto.Attributes(Propriedade) = value
                End Try

            End Set
        End Property

        ''' <summary>
        ''' Retorna SQL traduzido conforme retorno de MACROSUBSTSQL não prevendo a utilização de variáveis bind (tudo no texto).
        ''' </summary>
        ''' <param name="SQLRef">SQL a ser tratado, podendo o mesmo possuir referências :VARIAVEL, [:CAMPO] ou [:valor.CAMPO] para macrosubstituição.</param>
        ''' <param name="ListaDeOrigens">Lista de origens podendo ser LOGON, CONTAINER ou parâmetros nominais por exemplo: VARIAVEL, "CONTEUDO", :OUTRAVAR, 10.</param>
        ''' <returns>Retorna SQL tratado, com variáveis convertidas em texto.</returns>
        ''' <remarks></remarks>
        Shared Function MacroSubstSQLText(ByVal SQLRef As String, ByVal ParamArray ListaDeOrigens() As Object) As String
            Dim SQL As String = SQLRef
            MacroSubstSQL(SQL, Nothing, ListaDeOrigens)
            Return SQL
        End Function

        ''' <summary>
        ''' Obsoleta: Rotina de macrosubstituição de variáveis conforme fontes especificadas como lista de parâmetros, formulários, logon, etc. Esta função será descontinuada.
        ''' </summary>
        ''' <param name="SQLRef">Texto a ser tratado, onde será efetuada a busca por parâmetros como [:variavel].</param>
        ''' <param name="ListaOrigens">Origens possíveis para pesquisa das variáveis.</param>
        ''' <remarks></remarks>
        <Obsolete("Esta função será descontinuada.")> _
        Shared Sub MacroSubst(ByRef SQLRef As String, ByVal ParamArray ListaOrigens() As Object)
            ' se não tem campo sai imediatamente
            Dim _gex_expr As String = "\[:([^\]]*)\]"
            Dim resultG As Match = Regex.Match(SQLRef, _gex_expr)
            If resultG.Captures.Count = 0 Then
                Exit Sub
            End If

            ' identifica possíveis origens
            Dim Origens As ArrayList = ParamArrayToArrayList(ListaOrigens)

            ' seleciona origens complexas
            Dim OrigensObj As ArrayList = New ArrayList
            For Each Item As Object In Origens
                If Not TypeOf (Item) Is String Then
                    OrigensObj.Add(Item)
                End If
            Next


            ' procura os parâmetros
            While resultG.Captures.Count <> 0
                Dim Param As String = resultG.Groups(1).Value
                Dim TrocarPara As String = ":" & Param
                Dim Conteudo As Object = Nothing

                ' procura string
                Dim Pos As Integer = Origens.IndexOf(":" & Param)
                If Pos <> -1 Then
                    Conteudo = Origens.Item(Pos + 1)
                Else

                    ' procura nas origens especificadas
                    For Each Origem As Object In OrigensObj
                        Conteudo = OrigemParam(Origem, Param)
                        If Not IsNothing(Conteudo) Then
                            Exit For
                        End If
                    Next
                End If

                ' ocorre um erro caso conteúdo não seja encontrado
                If IsNothing(Conteudo) Then
                    Throw New Exception("Em MacroSubst, variável não identificada " & resultG.Value & ".")
                End If

                ' atualiza termo conforme parâmetro encontrado
                SQLRef = SQLRef.Substring(0, resultG.Index) & Conteudo & SQLRef.Substring(resultG.Index + resultG.Length)

                resultG = Regex.Match(SQLRef, _gex_expr)
            End While
        End Sub

        ''' <summary>
        ''' Troca em TEXTO variáveis a partir de fontes LOGON e CONTAINERS, podendo retornar em variáveis bind ou no próprio SQL (PARAMS = Nothing).
        ''' </summary>
        ''' <param name="SQLRef">Texto SQL contendo variáveis: SELECT [:valor.CAMPO], * FROM TABELA WHERE CAMPO=[:CAMPO]...</param>
        ''' <param name="Params">Lista de parâmetros conhecidos passados e retorno de variáveis bind para submissão ao interpretador SQL.</param>
        ''' <param name="ListaOrigens">Lista de origens podendo ser LOGON ou CONTAINERS, sendo a busca realizada por FINDGERAL (filhos e em seguida pais).</param>
        ''' <remarks></remarks>
        Shared Sub MacroSubstSQL(ByRef SQLRef As String, ByRef Params As ArrayList, ByVal ParamArray ListaOrigens() As Object)
            ' se não tem campo sai imediatamente
            If IsNothing(SQLRef) OrElse SQLRef = "" Then
                Exit Sub
            End If
            Dim _gex_expr As String = "\[:([^\]]*)\]"
            Dim resultG As Match = Regex.Match(SQLRef, _gex_expr)
            If resultG.Captures.Count = 0 Then
                Exit Sub
            End If

            ' identifica possíveis origens
            Dim Origens As ArrayList = ParamArrayToArrayList(ListaOrigens, Params)

            ' seleciona origens complexas
            Dim OrigensObj As ArrayList = New ArrayList
            For Each Item As Object In Origens
                If Not TypeOf (Item) Is String Then
                    OrigensObj.Add(Item)
                End If
            Next

            ' procura os parâmetros
            While resultG.Captures.Count <> 0
                Dim Param As String = resultG.Groups(1).Value
                Dim Tipo As String = "exprsql"
                ' vazio corresponde à exprsql
                ' possíveis valores... [:tipo.campo]
                '   TIPO = ...
                '    "[:campo]" ou  "[:exprsql.campo]", converte campo conforme tipo utilizando SQLEXPR, fazendo uso dos params ou caso vazio, na string sql;
                '    "[:valor.campo]", faz substituição diretamente no sql, colocando exatamente o conteúdo, utilizado para declaração de nomes etc...
                With Split(Param, ".")
                    If .Length > 1 Then
                        Tipo = .GetValue(0)
                        Param = .GetValue(1)
                    End If
                End With
                Dim TrocarPara As String = ":" & Param
                Dim Conteudo As Object = Nothing

                ' procura string
                Dim Pos As Integer = Origens.IndexOf(":" & Param)
                If Pos <> -1 Then
                    Conteudo = Origens.Item(Pos + 1)
                Else

                    ' procura nas origens especificadas
                    For Each Origem As Object In OrigensObj
                        Conteudo = OrigemParam(Origem, Param)
                        If Not IsNothing(NZ(Conteudo, Nothing)) Then
                            Exit For
                        End If
                    Next
                End If

                ' ocorre um erro caso conteúdo não seja encontrado
                If IsNothing(Conteudo) Then
                    Throw New Exception("Em MacroSubstSQL, variável não identificada " & resultG.Value & ".")
                End If

                ' atualiza termo conforme parâmetro encontrado
                If Compare(Tipo, "Valor") Then

                    ' valor direto sem aspas ou qq tratamento.
                    SQLRef = SQLRef.Substring(0, resultG.Index) & Conteudo & SQLRef.Substring(resultG.Index + resultG.Length)

                ElseIf IsNothing(Params) Then

                    ' valor direto com utilização sqlexpr
                    SQLRef = SQLRef.Substring(0, resultG.Index) & SqlExpr(Conteudo) & SQLRef.Substring(resultG.Index + resultG.Length)

                Else

                    ' atualiza conteúdo como parâmetro
                    Params.Add(TrocarPara)
                    Params.Add(Conteudo)
                    SQLRef = SQLRef.Substring(0, resultG.Index) & TrocarPara & SQLRef.Substring(resultG.Index + resultG.Length)
                End If

                resultG = Regex.Match(SQLRef, _gex_expr)
            End While
        End Sub

        ''' <summary>
        ''' Intepreta parâmetro solicitado com base nas possibilidades de pesquisa, sendo logon e continentes em geral as fontes.
        ''' </summary>
        ''' <param name="Origem">Origem de pesquisa.</param>
        ''' <param name="Param">Parâmetro a ser procurado na origem.</param>
        ''' <returns>Retorna o parâmetro como valor.</returns>
        ''' <remarks></remarks>
        Shared Function OrigemParam(ByVal Origem As Object, ByVal Param As String) As Object
            Try
                If Origem.GetType.ToString = "Icraft+LogonSession" Then
                    Return Prop(Origem, Param)
                Else
                    Dim Ctl As Control = Form.FindGeral(Origem, Param)
                    If Not IsNothing(Ctl) Then
                        Return Controle.ValorAtual(Ctl)
                    End If
                End If
            Catch
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' Coloca Conteudo no formato de uma expressão JavaScript Válida.
        ''' </summary>
        ''' <param name="Conteudo">O conteúdo a ser convertido.</param>
        ''' <returns>Retorna Conteudo convertido em uma expressão JavaScript válida.</returns>
        ''' <remarks></remarks>
        Shared Function JSExpr(ByVal Conteudo As String) As String
            Return "'" & Conteudo.Replace("'", "\'") & "'"
        End Function

        ''' <summary>
        ''' Retorna um conteúdo formatado de acordo com seu tipo para sua utilização em expressões SQL.
        ''' </summary>
        ''' <param name="Conteudo">Conteúdo a ser formatado.</param>
        ''' <returns>Texto para concatenação em expressões sql.</returns>
        ''' <remarks>Nem todos os tipos estão tratados. Serão configurados conforme a necessidade.</remarks>
        Shared Function SqlExpr(ByVal Conteudo As Object, Optional ByVal CaracAbreFechaString As String = "'") As String
            If TypeOf (Conteudo) Is String Then
                Return CaracAbreFechaString & Replace(Conteudo, CaracAbreFechaString, CaracAbreFechaString & CaracAbreFechaString) & CaracAbreFechaString
            ElseIf TypeOf (Conteudo) Is DBNull Then
                Return "NULL"
            ElseIf TypeOf Conteudo Is Decimal OrElse TypeOf Conteudo Is Double OrElse TypeOf Conteudo Is Single OrElse TypeOf Conteudo Is Int32 OrElse TypeOf Conteudo Is Byte Then
                Return Str(Conteudo)
            ElseIf TypeOf (Conteudo) Is Boolean Then
                Return IIf(Conteudo, Boolean.TrueString, Boolean.FalseString)
            ElseIf TypeOf (Conteudo) Is Date Then
                Return "#" & Format(Conteudo, "yyyy-MM-dd HH:mm:ss") & "#"
            Else
                Throw New Exception("Tipo desconhecido " & Conteudo.GetType.ToString & " para obtenção de expressão para sql.")
            End If
        End Function

        ''' <summary>
        ''' Retorna valor padrão se for Nothing, Nulo ou Vazio (ou zero no caso de tipo numérico).
        ''' </summary>
        ''' <param name="Valor">Valor a ser checado.</param>
        ''' <param name="Def">Default a ser retornado caso seja Nothing, Nulo ou vazio.</param>
        ''' <returns>Valor checado ou valor default caso Nothing, Nulo ou vazio (zero se o tipo for numérico).</returns>
        ''' <remarks></remarks>
        Shared Function NZV(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
            Dim Result As Object = NZ(Valor, Def)
            If TypeOf Result Is String AndAlso Result = "" Then
                Return Def
            ElseIf TypeOf Result Is Decimal AndAlso Result = 0 Then
                Return Def
            ElseIf TypeOf Result Is Double AndAlso Result = 0 Then
                Return Def
            ElseIf TypeOf Result Is Single AndAlso Result = 0 Then
                Return Def
            ElseIf TypeOf Result Is Int32 AndAlso Result = 0 Then
                Return Def
            ElseIf TypeOf Result Is Byte AndAlso Result = 0 Then
                Return Def
            End If
            Return Result
        End Function

        ''' <summary>
        ''' Caso o objeto inicial não exista (ismissing) ou seja nulo (dbnull), retorna o segundo parâmetro.
        ''' </summary>
        ''' <param name="Valor">Parâmetro a ser analisado.</param>
        ''' <param name="Def">Parâmetro default caso o primeiro parâmetro não exista ou seja nulo.</param>
        ''' <returns>Retorna primeiro parâmetro ou segundo caso o primeiro não exista ou seja nulo, sempre convertendo para o tipo do segundo parâmetro.</returns>
        ''' <remarks></remarks>
        Shared Function NZ(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
            Dim tipo As String

            If Not IsNothing(Def) Then
                tipo = Def.GetType.ToString
            ElseIf IsNothing(Valor) Then
                Return Nothing
            Else
                tipo = Valor.GetType.ToString
            End If

            If IsNothing(Valor) OrElse IsDBNull(Valor) OrElse ((tipo = "System.DateTime" Or Valor.GetType.ToString = "System.DateTime") AndAlso Valor = CDate(Nothing)) Then
                Valor = Def
            End If

            Select Case tipo
                Case "System.Decimal"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Decimal)
                    End If
                    Return CType(Valor, Decimal)
                Case "System.String"
                    If Valor.GetType.ToString = "System.Byte[]" Then
                        Return CType(ByteArrayToObject(Valor), String)
                    End If
                    If Valor.GetType.ToString = "Icraft.IcftBase+LogonSession" Then
                        Return CType(Valor, LogonSession).ToString
                    ElseIf Valor.GetType.IsEnum Then
                        Return Valor.ToString
                    End If
                    Return CType(Valor, String)
                Case "System.Double"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Double)
                    End If
                    Return CType(Valor, Double)
                Case "System.Boolean"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return False
                    End If
                    Return CType(Valor, Boolean)
                Case "System.DateTime"
                    Return CType(Valor, System.DateTime)
                Case "System.Single"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Single)
                    End If
                    Return CType(Valor, System.Single)
                Case "System.Byte"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Byte)
                    End If
                    Return CType(Valor, System.Byte)
                Case "System.Char"
                    Return CType(Valor, System.Char)
                Case "System.SByte"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, SByte)
                    End If
                    Return CType(Valor, System.SByte)
                Case "System.Int32"
                    If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                        Return CType(0, Int32)
                    End If
                    Return CType(Valor, Int32)
                Case "System.DBNull"
                    Return Valor
                Case "System.Collections.ArrayList"
                    Return ParamArrayToArrayList(Valor)
            End Select

            Return CType(Valor, String)
        End Function

        ''' <summary>
        ''' Transforma um paramarray em arraylist.
        ''' </summary>
        ''' <param name="PARAMS">Lista de parâmetros podendo ser um arraylist ou paramarray.</param>
        ''' <returns>Retornará um arraylist contendo a lista de parâmetros.</returns>
        ''' <remarks></remarks>
        Shared Function ParamArrayToArrayList(ByVal ParamArray Params() As Object) As Object

            ' caso não existam parâmetros
            If IsNothing(Params) OrElse Params.Length = 0 Then
                Return New ArrayList
            End If

            ' caso já seja um arraylist
            If Params.Length = 1 And TypeOf (Params(0)) Is ArrayList Then
                Return Params(0)
            End If

            ' caso tenha que juntar
            Dim ListaParametros As ArrayList = New ArrayList
            For Each Item As Object In Params
                If Not IsNothing(Item) Then

                    ' >> TIPOS PREVISTOS EM ARRAYLIST...
                    ' array
                    ' arraylist
                    ' string
                    ' dataset
                    ' datarowcollection

                    If TypeOf Item Is Array Then
                        For Each SubItem As Object In Item
                            ListaParametros.AddRange(ParamArrayToArrayList(SubItem))
                        Next
                    ElseIf TypeOf Item Is ArrayList Then
                        ListaParametros.AddRange(Item)
                    ElseIf TypeOf Item Is String Then
                        ListaParametros.Add(Item)
                    ElseIf TypeOf Item Is DataSet Then
                        For Each Row As DataRow In Item.Tables(0).rows
                            For Each Campo As Object In Row.ItemArray
                                ListaParametros.Add(Campo)
                            Next
                        Next
                    ElseIf TypeOf Item Is DataRow Then
                        For Each Campo As Object In CType(Item, DataRow).ItemArray
                            ListaParametros.Add(Campo)
                        Next
                    ElseIf TypeOf Item Is System.IO.FileInfo Then
                        ListaParametros.Add(Item.name)
                    Else
                        ListaParametros.Add(Item)
                    End If
                End If
            Next
            Return ListaParametros
        End Function

        ''' <summary>
        ''' Transforma array de bytes em objeto.
        ''' </summary>
        ''' <param name="Bytes">Array de bytes a ser transferida para objeto.</param>
        ''' <returns>Objeto criado a partir do array de bytes.</returns>
        ''' <remarks></remarks>
        Shared Function ByteArrayToObject(ByVal Bytes() As Byte) As Object
            Dim Obj As Object = Nothing
            Try
                Dim fs As System.IO.MemoryStream = New System.IO.MemoryStream
                Dim formatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                fs.Write(Bytes, 0, Bytes.Length)
                fs.Seek(0, IO.SeekOrigin.Begin)

                Obj = formatter.Deserialize(fs)
            Catch
            End Try
            Return Obj
        End Function

        ''' <summary>
        ''' Converte para outro tipo (ctype) especificado por uma string. 
        ''' </summary>
        ''' <param name="Obj">Objeto original.</param>
        ''' <param name="Tipo">String que especifica o tipo ex: System.String, System.Int32, etc.</param>
        ''' <returns>Retorna o objeto no tipo especificado.</returns>
        ''' <remarks></remarks>
        Shared Function CTypeStr(ByVal Obj As Object, ByVal Tipo As String) As Object
            If String.Compare(Tipo, "System.String", True) = 0 Or String.Compare(Tipo, "System.Byte[]", True) = 0 Then
                Return CType(Obj, String)
            ElseIf String.Compare(Tipo, "System.UInt32", True) = 0 Then
                Try
                    Return CType(Obj, UInt32)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para UInt32")
                End If
            ElseIf String.Compare(Tipo, "System.Int32", True) = 0 Then
                Try
                    Return CType(Obj, Int32)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para UInt32")
                End If
            ElseIf String.Compare(Tipo, "System.Int16", True) = 0 Then
                Try
                    Return CType(Obj, Int16)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para int16")
                End If
            ElseIf String.Compare(Tipo, "System.Int64", True) = 0 Then
                Try
                    Return CType(Obj, Int64)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para int64")
                End If
            ElseIf String.Compare(Tipo, "System.Integer", True) = 0 Then
                Try
                    Return CType(Obj, Integer)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para integer")
                End If
            ElseIf String.Compare(Tipo, "System.Boolean", True) = 0 Then
                Try
                    Return CType(Obj, Boolean)
                Catch
                End Try
                If TypeOf Obj Is String Then
                    If Obj = "on" Then
                        Return True
                    End If
                End If
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para boolean")
                End If
            ElseIf String.Compare(Tipo, "System.Date", True) = 0 Then
                Try
                    Return CType(Obj, Date)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para date")
                End If
            ElseIf String.Compare(Tipo, "System.DateTime", True) = 0 Then
                Try
                    Return CType(Obj, DateTime)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para datetime")
                End If
            ElseIf String.Compare(Tipo, "System.Decimal", True) = 0 Then
                Try
                    Return CType(Obj, Decimal)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para decimal")
                End If
            ElseIf String.Compare(Tipo, "System.Double", True) = 0 Then
                Try
                    Return CType(Obj, Double)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para Double")
                End If
            ElseIf String.Compare(Tipo, "System.Single", True) = 0 Then
                Try
                    Return CType(Obj, Single)
                Catch
                End Try
                If CType(Obj, String).Trim = "" Then
                    Return Convert.DBNull
                Else
                    Err.Raise(20000, , "ctypestr falhou ao converter obj para Single")
                End If
            ElseIf IsNothing(Tipo) Then
                Return ""
            Else
                Err.Raise(20000, , "ctypestr tipo não previsto: " & Tipo)
            End If
            Return Obj
        End Function

        ''' <summary>
        ''' Extrai atributos de registros ou items do objeto. Utilizada para obter campo de um determinado dataset ou dataview.
        ''' </summary>
        ''' <param name="Obj">Objeto a ser explorado como por exemplo: Dataset.Tables(0).Rows.</param>
        ''' <param name="PropRel">Propriedade relacionada existente em todos os registros como nome ou número de um campo.</param>
        ''' <returns>Retorna o campo de cada registro como um índice de ArrayList.</returns>
        ''' <remarks>Exemplo: ItemsToArrayList(Dataset.Tables(0).Rows, "MeuCampo")</remarks>
        Shared Function ItemsToArrayList(ByVal Obj As Object, ByVal PropRel As Object) As ArrayList
            Dim Lista As ArrayList = New ArrayList
            For Each Item As Object In Obj
                Try
                    Lista.Add(Atrib(Item, PropRel))
                Catch
                    Try
                        Lista.Add(Item(PropRel))
                    Catch
                        Lista.Add(Item.Attributes(PropRel))
                    End Try
                End Try
            Next
            Return Lista
        End Function

        ''' <summary>
        ''' Obtém itens de um objeto considerando uma propriedade específica.
        ''' </summary>
        ''' <param name="Obj">Objeto que contém a lista de itens (rows por exemplo).</param>
        ''' <param name="PropRel">Nome da propriedade ou campo a ser obtido para cada linha.</param>
        ''' <returns>Retorna uma lista, que poderá ser atribuída ao datasource para preenchimento automático de combox.</returns>
        ''' <remarks></remarks>
        Shared Function ItemsToObject(ByVal Obj As Object, ByVal PropRel As Object) As List(Of Object)
            Dim Lista As List(Of Object) = New List(Of Object)
            For Each Item As Object In Obj
                Try
                    Lista.Add(Atrib(Item, PropRel))
                Catch
                    Try
                        Lista.Add(Item(PropRel))
                    Catch
                        Lista.Add(Item.Attributes(PropRel))
                    End Try
                End Try
            Next
            Return Lista
        End Function


        ''' <summary>
        ''' Tratamento de idioma padronizado para condicionamento de ambiente.
        ''' </summary>
        ''' <remarks></remarks>
        Class Idioma
            ''' <summary>
            ''' Idioma definido para o ambiente.
            ''' </summary>
            ''' <param name="Page">Página a ser avaliada.</param>
            ''' <value>Valor TipoIdioma a ser definido para contexto do ambiente.</value>
            ''' <returns>Valor definido TipoIdioma no contexto do ambiente.</returns>
            ''' <remarks></remarks>
            Shared Property DoAmbiente(ByVal Page As Page) As TipoIdioma
                Get
                    Dim Termo As String = Page.Session("IdiomaDoAmbiente")
                    If Not IsNothing(Termo) Then
                        Return CType(Termo, TipoIdioma)
                    End If
                    Return TipoIdioma.PT_BR
                End Get
                Set(ByVal value As TipoIdioma)
                    Page.Session("IdiomaDoAmbiente") = value
                End Set
            End Property

            ''' <summary>
            ''' Verifica especificação da página para definição de contexto do ambiente.
            ''' </summary>
            ''' <param name="Page">Página a ser tratada.</param>
            ''' <remarks></remarks>
            Shared Sub Verifica(ByVal Page As Page)
                If Page.Request.Form("__EVENTTARGET") = "DEFINE_IDIOMA" Then
                    Dim Param As String = Page.Request.Form("__EVENTARGUMENT")
                    If Compare(Param, "EN") Then
                        DoAmbiente(Page) = TipoIdioma.EN
                    ElseIf Compare(Param, "ES") Then
                        DoAmbiente(Page) = TipoIdioma.ES
                    ElseIf Compare(Param, "PT_BR") Then
                        DoAmbiente(Page) = TipoIdioma.PT_BR
                    End If
                End If
            End Sub
        End Class

        ''' <summary>
        ''' Classe para registro de logon de usuário e variáveis relacionadas.
        ''' </summary>
        ''' <remarks></remarks>
        Class LogonSession
            Private _id As String = Nothing
            Private _usuario As String = Nothing
            Private _momento As Date = Nothing
            Private _site As String = Nothing
            Private _senha As String = Nothing
            Private _outros As New ArrayList

            Public Shadows Function ToString() As String
                Dim txt As New StringBuilder
                txt.Append("LogonSession(")
                txt.Append("id=" & NZ(_id, "") & ";")
                txt.Append("_usuario=" & NZ(_usuario, "") & ";")
                txt.Append("_momento=" & Format(NZV(_momento, Nothing), "dd/MM/yyyy HH:mm:ss") & ";")
                For z As Integer = 0 To _outros.Count - 1 Step 2
                    txt.Append(_outros(z) & "=")
                    txt.Append(NZ(_outros(z + 1), ""))
                    txt.Append(";")
                Next
                txt.Append("_site=" & NZ(_site, ""))
                txt.Append(")")
                Return txt.ToString
            End Function


            ''' <summary>
            ''' Identificação para armazenamento de logon do tipo 'GERAL' ou algum específico para múltiplos logons.
            ''' </summary>
            ''' <value>Especificação do tipo de logon.</value>
            ''' <returns>Especificação do tipo de logon.</returns>
            ''' <remarks></remarks>
            Public Property Id() As String
                Get
                    Return _id
                End Get
                Set(ByVal value As String)
                    _id = value
                End Set
            End Property

            ''' <summary>
            ''' Usuário que efetuou logon.
            ''' </summary>
            ''' <value>Login do usuário que efetuou logon.</value>
            ''' <returns>Login do usuário que efetou logon.</returns>
            ''' <remarks></remarks>
            Public Property Usuario() As String
                Get
                    Return _usuario
                End Get
                Set(ByVal value As String)
                    _usuario = value
                End Set
            End Property

            ''' <summary>
            ''' Momento de logon.
            ''' </summary>
            ''' <value>Momento (data e hora) de logon.</value>
            ''' <returns>Momento (data e hora) de logon.</returns>
            ''' <remarks></remarks>
            Public Property Momento() As Date
                Get
                    Return _momento
                End Get
                Set(ByVal value As Date)
                    _momento = value
                End Set
            End Property

            ''' <summary>
            ''' Nome do site.
            ''' </summary>
            ''' <value></value>
            ''' <returns>Nome do site.</returns>
            ''' <remarks>Nome do site.</remarks>
            Public Property Site() As String
                Get
                    Return _site
                End Get
                Set(ByVal value As String)
                    _site = value
                End Set
            End Property

            ''' <summary>
            ''' Senha de acesso.
            ''' </summary>
            ''' <value>Senha de acesso.</value>
            ''' <returns>Senha de acesso.</returns>
            ''' <remarks></remarks>
            Public Property Senha() As String
                Get
                    Return _senha
                End Get
                Set(ByVal value As String)
                    _senha = value
                End Set
            End Property

            ''' <summary>
            ''' Outras propriedades a serem armazenadas pelo Logon.
            ''' </summary>
            ''' <param name="Propriedade">Nome da propriedade.</param>
            ''' <value>Valor da propriedade.</value>
            ''' <returns>Valor da propriedade armazenada.</returns>
            ''' <remarks></remarks>
            Public Property ExtendedProps(ByVal Propriedade As String) As Object
                Get
                    Dim Pos As Integer = _outros.IndexOf(":" & Propriedade)
                    If Pos >= 0 Then
                        Return _outros(Pos + 1)
                    End If
                    Return Nothing
                End Get
                Set(ByVal value As Object)
                    Dim Pos As Integer = _outros.IndexOf(":" & Propriedade)
                    If Pos >= 0 Then
                        _outros(Pos + 1) = value
                        Exit Property
                    End If
                    _outros.Add(":" & Propriedade)
                    _outros.Add(value)
                End Set
            End Property

            ''' <summary>
            ''' Acesso aos atributos e propriedades expandidas.
            ''' </summary>
            ''' <param name="Nome">Nome da propriedade tratada.</param>
            ''' <value>Valor da propriedade tratada.</value>
            ''' <returns>Valor da propriedade solicitada.</returns>
            ''' <remarks></remarks>
            Default Property Attributes(ByVal Nome As String) As String
                Get
                    If Compare(Nome, "Id") Then
                        Return _id
                    ElseIf Compare(Nome, "Usuario") Then
                        Return _usuario
                    ElseIf Compare(Nome, "Momento") Then
                        Return _momento
                    ElseIf Compare(Nome, "Site") Then
                        Return _site
                    ElseIf Compare(Nome, "Senha") Then
                        Return _senha
                    Else
                        Dim Prop As Object = ExtendedProps(Nome)
                        If IsNothing(Prop) Then
                            Throw New Exception("Em Attributes de Logon, atributo '" & Nome & "' inválido para objeto " & Me.GetType.ToString & ".")
                        Else
                            Return Prop
                        End If
                    End If
                    Return Nothing
                End Get

                Set(ByVal value As String)
                    If Compare(Nome, "Id") Then
                        _id = value
                    ElseIf Compare(Nome, "Usuario") Then
                        _usuario = value
                    ElseIf Compare(Nome, "Momento") Then
                        _momento = value
                    ElseIf Compare(Nome, "Site") Then
                        _site = value
                    ElseIf Compare(Nome, "Senha") Then
                        _senha = value
                    Else
                        Throw New Exception("Em Attributes de Logon, atributo " & value & " inválido para objeto " & Me.GetType.ToString & ".")
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Criação de login para registro de acesso de usuário.
            ''' </summary>
            ''' <param name="Pagina">Página na qual é efetuado o login.</param>
            ''' <param name="Usuario">Usuário que efetua acesso.</param>
            ''' <param name="Senha">Senha do usuário.</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal Pagina As Page, ByVal Usuario As String, ByVal Senha As String)
                ' cria chave com area e usuario
                _id = Pagina.Session.SessionID
                _usuario = Usuario
                _momento = Now
                _site = WebConf("site_nome")
                _senha = Senha
            End Sub

        End Class

        ''' <summary>
        ''' Classe para tratamento de form icraft.
        ''' </summary>
        ''' <remarks></remarks>
        Public Class Form

            ''' <summary>
            ''' Carrega definições obtidas através do gerador.
            ''' </summary>
            ''' <param name="ds">Lista de definições de campos que será retornada.</param>
            ''' <param name="StrGerador">String de conexão com gerador.</param>
            ''' <param name="Sistema">Nome do sistema.</param>
            ''' <param name="Tabela">Nome da tabela desejada.</param>
            ''' <param name="Params">Parâmetros de filtro.</param>
            ''' <remarks></remarks>
            Shared Sub CarregaDef(ByRef ds As DataSet, ByVal StrGerador As String, ByVal Sistema As String, ByVal Tabela As String, ByVal ParamArray Params() As Object)
                If StrGerador <> "" And Sistema <> "" And Tabela <> "" Then
                    Dim Pos As Integer = InStr(Tabela, ".")
                    If Pos <> 0 Then
                        Tabela = Mid(Tabela, Pos + 1)
                    End If

                    Dim dsgera As DataSet = DSCarrega("SELECT CAMPO,ETIQ,DESCR,PROP_EXTEND,FORMATO,TIPO_ORACLE,TIPO_ACCESS,TIPO_MYSQL,AUTO,VALOR_PADRAO FROM GER_CAMPO WHERE SISTEMA=:SISTEMA AND TABELA=:TABELA", StrGerador, ":SISTEMA", Sistema, ":TABELA", Tabela, Params)
                    If dsgera.Tables.Count > 0 AndAlso dsgera.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dsgera.Tables(0).Rows
                            Pos = ds.Tables(0).Columns.IndexOf(row("CAMPO"))
                            If Pos <> -1 Then
                                ds.Tables(0).Columns(Pos).ExtendedProperties("Etiq") = NZ(row("ETIQ"), "")
                                ds.Tables(0).Columns(Pos).ExtendedProperties("Descr") = NZ(row("DESCR"), "")
                                ds.Tables(0).Columns(Pos).ExtendedProperties("Props") = NZ(row("PROP_EXTEND"), "")
                                ds.Tables(0).Columns(Pos).ExtendedProperties("Formato") = NZ(row("FORMATO"), "")
                                ds.Tables(0).Columns(Pos).ExtendedProperties("Auto") = NZ(row("AUTO"), "")
                                ds.Tables(0).Columns(Pos).ExtendedProperties("ValorPadrao") = NZ(row("VALOR_PADRAO"), "")

                                Dim Tipo As String = Microsoft.VisualBasic.Switch(NZ(row("TIPO_ORACLE"), "") <> "", row("TIPO_ORACLE"), NZ(row("TIPO_ACCESS"), "") <> "", row("TIPO_ACCESS"), True, row("TIPO_MYSQL"))
                                Dim Tam As Integer = 0
                                If Tipo <> "" Then
                                    ds.Tables(0).Columns(Pos).ExtendedProperties("Tamanho") = RegexGroup(Tipo, ".*(VARCHAR2|VARCHAR|CHAR|TEXT).*\((.*)\)", 2).Value
                                Else
                                    ds.Tables(0).Columns(Pos).ExtendedProperties("Tamanho") = ""
                                End If
                            End If
                        Next
                    End If
                End If
            End Sub

            ''' <summary>
            ''' Uma das formas de carga de registro, considerando seu conteúdo.
            ''' </summary>
            ''' <param name="ListaControles">Lista de controles a serem tratados.</param>
            ''' <param name="Prefixo">Prefixo dos campos a serem preenchidos.</param>
            ''' <param name="Registros">Dataset contendo registros a serem pesquisados.</param>
            ''' <param name="Registro">Número do registro a ser apresentado.</param>
            ''' <param name="SomenteEstrut">Obtenção somente da estrutura.</param>
            ''' <param name="Params">Parâmetros para filtro das opções de registro.</param>
            ''' <remarks></remarks>
            Overloads Shared Sub CarregaReg(ByRef ListaControles As Object, ByVal Prefixo As String, ByRef Registros As Object, ByRef Registro As Object, ByVal SomenteEstrut As Boolean, ByVal ParamArray Params() As Object)
                If Not Registros Is Nothing Then

                    ' trata origem
                    Dim DV As DataView = Nothing

                    If Registros.GetType.ToString = "System.Data.DataSet" AndAlso CType(Registros, DataSet).Tables.Count = 1 Then
                        DV = CType(Registros, DataSet).Tables(0).DefaultView
                    ElseIf Registros.GetType.ToString = "System.Data.DataView" Then
                        DV = CType(Registros, DataView)
                    End If

                    ' busca registro
                    Dim Row As DataRowView = Nothing
                    If Not IsNothing(Registro) Then
                        If Registro.GetType.ToString = "System.Data.DataRowView" Then
                            Row = CType(Registro, DataRowView)
                        ElseIf Registro.GetType.ToString = "System.Int32" AndAlso Not IsNothing(DV) AndAlso DV.Count > CType(Registro, Integer) Then
                            Row = DV.Item(CType(Registro, Integer))
                        End If
                    End If

                    ' atualiza controles na tela
                    For Each Ctl As Control In Controles(ListaControles, Prefixo)
                        Dim NomeControle As String = Ctl.ID.Substring(Len(Prefixo))
                        If Not IsNothing(DV) And Controle.Tipo(Ctl) = "" Then
                            Controle.Tipo(Ctl) = DV.Table.Columns(NomeControle).DataType.ToString
                        End If
                        If Not SomenteEstrut Then
                            If IsNothing(Row) Then
                                Controle.ValorAtual(Ctl, True) = ""
                                Controle.ValorAnterior(Ctl) = Controle.ValorAtual(Ctl)
                            Else
                                Dim Formato As String = Controle.Mascara(Ctl)
                                Controle.ValorAtual(Ctl) = Row(NomeControle)
                                Controle.ValorAnterior(Ctl) = Row(NomeControle)
                            End If
                        End If
                    Next

                End If
            End Sub

            ''' <summary>
            ''' Uma das opções de carga de registros em campos.
            ''' </summary>
            ''' <param name="ListaControles">Lista de controles.</param>
            ''' <param name="Prefixo">Prefixo dos controles desejados.</param>
            ''' <param name="SelectOuTabela">Select ou nome da tabela a ser pesquisa.</param>
            ''' <param name="Filtro">Filtro para composição do SQL de pesquisa.</param>
            ''' <param name="StrConn">Conexão para pesquisa das informações.</param>
            ''' <param name="Params">Parâmetros de filtro para pesquisa.</param>
            ''' <returns>Dataset obtido mediante SQL resultante do SelectouTabela, Filtro e Params.</returns>
            ''' <remarks></remarks>
            Overloads Shared Function CarregaReg(ByVal ListaControles As Object, ByVal Prefixo As String, ByVal SelectOuTabela As String, ByVal Filtro As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object) As DataSet
                Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
                Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)
                Dim SQL As String
                If Not SelectOuTabela.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase) Then
                    SQL = ""
                    For Each campo As String In Campos(ListaControles, Prefixo)
                        SQL &= IIf(SQL <> "", ", ", "") & campo
                    Next
                    SQL = "SELECT " & SQL & " FROM " & SelectOuTabela & IIf(Filtro <> "", " WHERE " & Filtro, "")
                Else
                    SQL = SelectOuTabela & IIf(Filtro <> "", " WHERE " & Filtro, "")
                End If

                Dim ds As DataSet = DSCarrega(SQL, ConnW, ListaParametros)
                If ds.Tables(0).Rows.Count = 0 Then
                    ds = DSCarregaEstrut(SQL, ConnW)
                ElseIf ds.Tables(0).Rows.Count > 1 Then
                    Dim sqlerr As String = ""
                    For Each row As DataRow In ds.Tables(0).Rows
                        Try
                            sqlerr &= vbCrLf
                            sqlerr &= row(1) & " "
                            sqlerr &= row(2) & " "
                            sqlerr &= row(3) & " "
                        Catch
                        End Try
                    Next
                    Err.Raise(20000, "Icraft.CarregaReg", "edição de registro único chamada com recordset contendo mais uma linha" & sqlerr)
                End If

                CarregaReg(ListaControles, Prefixo, ds, 0, False)
                Return ds
            End Function

            ''' <summary>
            ''' Exclusão de registro.
            ''' </summary>
            ''' <param name="ListaControles">Lista de controles.</param>
            ''' <param name="Prefixo">Prefixo dos controles considerados.</param>
            ''' <param name="DeleteOuTabela">SQL de deleção ou tabela.</param>
            ''' <param name="Filtro">Filtro adicional caso exista.</param>
            ''' <param name="StrConn">Nome da conexão.</param>
            ''' <param name="Params">Parâmetros para filtro.</param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Shared Function ExcluiReg(ByVal ListaControles As Object, ByVal Prefixo As String, ByVal DeleteOuTabela As String, ByVal Filtro As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object) As Integer
                Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
                Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

                ' pega string a partir do FROM
                Dim SQLTot As String = ExprSQL(DeleteOuTabela, "", Filtro, , ExprSQLTipo.Sel)
                Dim SQLDel As String = ExprSQL(DeleteOuTabela, "", Filtro, , ExprSQLTipo.Del)
                Dim TOT As Integer = DSValor("COUNT(*)", SQLTot, ConnW, "", ListaParametros)
                DSGrava(SQLDel, ConnW, ListaParametros)
                Return TOT - DSValor("COUNT(*)", SQLTot, ConnW, "", ListaParametros)
            End Function

            ''' <summary>
            ''' Gravação de registro para formulário.
            ''' </summary>
            ''' <param name="ListaControles">Opções de campos a serem pesquisados.</param>
            ''' <param name="Prefixo">Prefixo dos campos a serem pesquisados.</param>
            ''' <param name="SqlOuTabela">SQL ou tabela onde ocorrerá o registro.</param>
            ''' <param name="Filtro">Filtro adicional caso necessário.</param>
            ''' <param name="StrConn">Nome de conexão para registro.</param>
            ''' <param name="Params">Parâmetros adicionais para registro.</param>
            ''' <returns>TRUE caso registro ocorra satisfatoriamente ou FALSE caso contrário.</returns>
            ''' <remarks></remarks>
            Shared Function GravaReg(ByVal ListaControles As Object, ByVal Prefixo As String, ByVal SqlOuTabela As String, ByVal Filtro As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object) As Boolean
                Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
                Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

                Dim SQL As String = ""
                Dim UPD_SETS As String = "", INS_CAMPOS As String = "", INS_VALS As String = ""

                If Not (SqlOuTabela.StartsWith("INSERT", StringComparison.OrdinalIgnoreCase) Or SqlOuTabela.StartsWith("UPDATE", StringComparison.OrdinalIgnoreCase) Or SqlOuTabela.StartsWith("DELETE", StringComparison.OrdinalIgnoreCase)) Then

                    ' caso não seja sql, é tabela e precisamos complementar com campos
                    For Each ctl As Control In Controles(ListaControles, Prefixo)
                        Dim NomeControle As String = ctl.ID.Substring(Len(Prefixo))
                        If Not ListaParametros.Contains(":" & NomeControle) Then
                            ListaParametros.Add(":" & NomeControle)
                            ListaParametros.Add(Controle.ValorAtual(ctl))

                            ' prepara para qualquer operação padrão em tabela insert update ou delete
                            UPD_SETS &= IIf(UPD_SETS <> "", ", ", "") & NomeControle & " = :" & NomeControle
                        End If
                        INS_CAMPOS &= IIf(INS_CAMPOS <> "", ", ", "") & NomeControle
                        INS_VALS &= IIf(INS_VALS <> "", ", ", "") & ":" & NomeControle
                    Next

                    ' se não exista, inclui. caso contrário, altera
                    If DSValor("COUNT(*)", SqlOuTabela, ConnW, Filtro, Params) > 0 Then
                        SQL = "UPDATE " & SqlOuTabela & " SET " & UPD_SETS & " " & IIf(Filtro <> "", " WHERE " & Filtro, "")
                    Else
                        SQL = "INSERT INTO " & SqlOuTabela & " (" & INS_CAMPOS & ") VALUES (" & INS_VALS & ")"
                    End If
                Else
                    SQL = SqlOuTabela & IIf(Filtro <> "", " WHERE " & Filtro, "")
                End If
                DSGrava(SQL, ConnW, ListaParametros)
                Return True
            End Function

            ''' <summary>
            ''' Inclusão de registro com base em formulário.
            ''' </summary>
            ''' <param name="ListaControles">Opções de campos a serem pesquisados.</param>
            ''' <param name="Prefixo">Prefixo dos campos a serem pesquisados.</param>
            ''' <param name="SqlOuTabela">SQL ou tabela onde ocorrerá o registro.</param>
            ''' <param name="Filtro">Filtro adicional caso necessário.</param>
            ''' <param name="StrConn">Nome de conexão para registro.</param>
            ''' <param name="Params">Parâmetros adicionais para registro.</param>
            ''' <returns>TRUE caso registro ocorra satisfatoriamente ou FALSE caso contrário.</returns>
            ''' <remarks></remarks>
            Shared Function IncluiReg(ByVal ListaControles As Object, ByVal Prefixo As String, ByVal SqlOuTabela As String, ByVal Filtro As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object) As Boolean
                Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
                Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

                Dim SQL As String = ""
                Dim INS_CAMPOS As String = "", INS_VALS As String = ""

                ' caso não seja sql, é tabela e precisamos complementar com campos
                For Each ctl As Control In Controles(ListaControles, Prefixo)
                    Dim NomeControle As String = ctl.ID.Substring(Len(Prefixo))
                    If Not ListaParametros.Contains(":" & NomeControle) Then
                        ListaParametros.Add(":" & NomeControle)
                        ListaParametros.Add(Controle.ValorAtual(ctl))

                        ' prepara para qualquer operação padrão em tabela insert update ou delete
                    End If
                    INS_CAMPOS &= IIf(INS_CAMPOS <> "", ", ", "") & NomeControle
                    INS_VALS &= IIf(INS_VALS <> "", ", ", "") & ":" & NomeControle
                Next

                SQL = "INSERT INTO " & SqlOuTabela & " (" & INS_CAMPOS & ") VALUES (" & INS_VALS & ")"
                DSGrava(SQL, ConnW, ListaParametros)
                Return True
            End Function

            ''' <summary>
            ''' Alteração de registro com base em formulário.
            ''' </summary>
            ''' <param name="ListaControles">Opções de campos a serem pesquisados.</param>
            ''' <param name="Prefixo">Prefixo dos campos a serem pesquisados.</param>
            ''' <param name="SqlOuTabela">SQL ou tabela onde ocorrerá o registro.</param>
            ''' <param name="Chave">Especificação de chave anterior para troca.</param>
            ''' <param name="Filtro">Filtro adicional caso necessário.</param>
            ''' <param name="StrConn">Nome de conexão para registro.</param>
            ''' <param name="Params">Parâmetros adicionais para registro.</param>
            ''' <returns>TRUE caso registro ocorra satisfatoriamente ou FALSE caso contrário.</returns>
            ''' <remarks></remarks>
            Shared Function AlteraReg(ByVal ListaControles As Object, ByVal Prefixo As String, ByVal SqlOuTabela As String, ByVal Chave As String, ByVal Filtro As String, ByVal StrConn As Object, ByVal ParamArray Params() As Object) As Boolean
                Dim ListaParametros As ArrayList = ParamArrayToArrayList(Params)
                Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ListaParametros)

                Dim SQL As String = ""
                Dim UPD_SETS As String = ""

                Dim NovoFiltro As String = ""

                ' caso não seja sql, é tabela e precisamos complementar com campos
                For Each ctl As Control In Controles(ListaControles, Prefixo)
                    Dim NomeControle As String = ctl.ID.Substring(Len(Prefixo))
                    If Array.IndexOf(Chave.Split(";"), NomeControle) <> -1 Then
                        If Controle.ValorAtual(ctl) <> Controle.ValorAnterior(ctl) Then
                            ' mudou chave
                            NovoFiltro &= IIf(NovoFiltro <> "", " AND ", "") & NomeControle & " = :" & NomeControle
                            UPD_SETS &= IIf(UPD_SETS <> "", ", ", "") & NomeControle & " = :" & NomeControle
                            ListaParametros.Add(":" & NomeControle)
                            ListaParametros.Add(Controle.ValorAtual(ctl))
                        End If
                    Else
                        If Not ListaParametros.Contains(":" & NomeControle) Then
                            UPD_SETS &= IIf(UPD_SETS <> "", ", ", "") & NomeControle & " = :" & NomeControle
                            ListaParametros.Add(":" & NomeControle)
                            ListaParametros.Add(Controle.ValorAtual(ctl))
                        End If
                    End If
                Next

                SQL = "UPDATE " & SqlOuTabela & " SET " & UPD_SETS & " " & IIf(Filtro <> "", " WHERE " & Filtro, "")
                DSGrava(SQL, ConnW, ListaParametros)
                Return True
            End Function

            Shared Function Controles(ByVal Container As Object, ByVal Prefixos As String, ByRef JaVerificado As ArrayList) As Object
                Dim Lista As New ArrayList
                If IsNothing(JaVerificado) Then
                    JaVerificado = New ArrayList
                End If
                If Not TemNaLista(JaVerificado, Container.UniqueID) Then
                    JaVerificado.Add(Container.UniqueId)
                    Dim Id As String = NZ(Container.ID, "")

                    Dim Achou As Boolean = False
                    For Each Item As String In Split(Prefixos, ";")
                        If Id.StartsWith(Item, StringComparison.OrdinalIgnoreCase) Then
                            Achou = True
                            Exit For
                        End If
                    Next
                    If Achou Then
                        Lista.Add(Container)
                    End If
                    Dim SubControls As Integer = 0
                    Try
                        SubControls = Container.controls.count
                    Catch
                    End Try
                    If SubControls > 0 Then
                        For Each Ctl As Object In Container.Controls
                            If Not TemNaLista(JaVerificado, Ctl.UNIQUEID) Then
                                CopiaItens(Lista, Controles(Ctl, Prefixos, JaVerificado))
                            End If
                        Next
                    End If
                End If
                Return Lista
            End Function

            Shared Function Controles(ByVal Container As Object, ByVal Prefixo As String) As Object
                Dim Lista As New ArrayList
                For Each Opcoes As Object In Containers(Container)
                    Dim Opc As Object = Nothing
                    If TypeOf Opcoes Is Web.UI.ControlCollection Or TypeOf Opcoes Is Windows.Forms.Form.ControlCollection Then
                        Opc = Opcoes
                    Else
                        Opc = Opcoes.Controls
                    End If
                    For Each Controle As Object In Opc
                        If NZ(Prop(Controle, "ID"), "").StartsWith(Prefixo, StringComparison.OrdinalIgnoreCase) Then
                            Lista.Add(Controle)
                        End If
                    Next
                Next
                Return Lista
            End Function

            ''' <summary>
            ''' Conteúdo dos controles para registro em texto em algum local.
            ''' </summary>
            ''' <param name="Container">Container onde está os controles.</param>
            ''' <param name="Prefixo">Prefixo a ser pesquisado.</param>
            ''' <param name="Atribs">Atributos adicionais.</param>
            ''' <value>Definição de conteúdo a ser atribuída aos controles.</value>
            ''' <returns>Definição de conteúdo obtida a partir dos controles.</returns>
            ''' <remarks></remarks>
            Shared Property Conteudo(ByVal Container As Object, ByVal Prefixo As String, Optional ByVal Atribs As String = "") As Object
                Get
                    Dim Elems As ElementosStr
                    Elems = New ElementosStr("")
                    For Each ctl As Object In Form.Controles(Container, Prefixo)
                        Elems.Items(-1) = New ElementoStr(Prop(ctl, "ID") & ":" & ItemEncode(Prop(ctl)))
                        If Atribs <> "" Then
                            For Each Atrib As String In Split(Atribs, ";")
                                Elems.Items(-1) = New ElementoStr(Prop(ctl, "ID") & "|" & Atrib & ":" & ItemEncode(Prop(ctl, Atrib)))
                            Next
                        End If
                    Next
                    Return Elems.ToString
                End Get
                Set(ByVal value As Object)
                    Dim Elems As ElementosStr
                    Elems = New ElementosStr(NZ(value, ""))
                    For Each ctl As Object In Form.Controles(Container, Prefixo)
                        Prop(ctl) = ItemDecode(Elems(Prop(ctl, "ID")).Conteudo)
                        If Atribs <> "" Then
                            For Each Atrib As String In Split(Atribs, ";")
                                Dim conteudo As Object = Elems(Prop(ctl, "ID") & "|" & Atrib).Conteudo
                                Prop(ctl, Atrib) = ItemDecode(conteudo)
                            Next
                        End If
                    Next
                End Set
            End Property

            ''' <summary>
            ''' Busca Controle a partir do container informado, procurando também em todos os filhos. Utilizar FINDCONTROLESPECIAL para encontrar controle também em Paineis.
            ''' </summary>
            ''' <param name="Container">Objeto a partir do qual a busca será iniciada.</param>
            ''' <param name="Nome">Nome do contrle a ser procurado.</param>
            ''' <returns>Retorna o controle caso seja encontrado entre as dependências do CONTAINER ou NOTHING se o contrário.</returns>
            ''' <remarks>Utilizar FINDGERAL para buscar entre os filhos e pais do container.</remarks>
            Shared Function FindControl(ByVal Container As Object, ByVal Nome As String, Optional ByVal NaoProcurarEm As ArrayList = Nothing) As Control
                For Each Opcao As Object In Containers(Container)
                    ' findcontrol comum não encontrava itens em paineis
                    ' troquei por findcontrolespecial <<

                    If Not IsNothing(NaoProcurarEm) AndAlso NaoProcurarEm.Contains(NaoProcurarEm) Then
                        ' se existe a lista e obj está nesta lista, não procurar neste objeto
                    Else
                        Dim Ctl As Control = FindControlEspecial(Opcao, Nome)
                        If Not IsNothing(Ctl) Then
                            Return Ctl
                        End If
                    End If
                Next
                Return Nothing
            End Function

            ''' <summary>
            ''' Busca controle tanto nos filhos como nos pais. Começa procurando entre os filhos e, depois, vai subindo até o topo.
            ''' </summary>
            ''' <param name="Container">Objeto a partir do qual a procura se inicia. Filhos deste terão a prioridade.</param>
            ''' <param name="Nome">Nome do controle a ser procurado.</param>
            ''' <param name="NaoProcurarEm">Lista negra. Objetos que estiverem esta lista serão evitados.</param>
            ''' <returns>Retorna controle ou NOTHING caso a busca não tenha sucesso.</returns>
            ''' <remarks></remarks>
            Shared Function FindGeral(ByVal Container As Object, ByVal Nome As String, Optional ByVal NaoProcurarEm As ArrayList = Nothing) As Object
                Dim Obj As Object = Container
                Dim JaProcurados As ArrayList = New ArrayList
                Do While Not IsNothing(Obj)
                    Dim Ctl As Control = Form.FindControl(Obj, Nome, JaProcurados)
                    If Not IsNothing(Ctl) Then
                        Return Ctl
                    End If
                    JaProcurados.Add(Obj)
                    Obj = Obj.Parent
                Loop
                Return Nothing
            End Function

            Shared Function BuscaTipo(ByVal Raiz As Object, ByVal ParamArray Tipos() As Object) As Object
                Return BuscaTipo(Raiz, False, Nothing, Tipos)
            End Function

            Shared Function BuscaPrimeiroTipo(ByVal Raiz As Object, ByVal ParamArray Tipos() As Object) As Object
                Dim Obj As Object = BuscaTipo(Raiz, True, Nothing, Tipos)
                If Obj.count > 0 Then
                    Return Obj(0)
                End If
                Return Nothing
            End Function

            Shared Function BuscaTipo(ByVal Raiz As Object, ByVal PararNoPrimeiro As Boolean, ByRef JaVerificado As ArrayList, ByVal ParamArray Tipos() As Object) As Object
                Dim TiposG As ArrayList = ParamArrayToArrayList(Tipos)
                Dim Lista As New ArrayList
                If IsNothing(JaVerificado) Then
                    JaVerificado = New ArrayList
                End If

                If Not TemNaLista(JaVerificado, Raiz.UNIQUEID) Then
                    JaVerificado.Add(Raiz.UNIQUEID)
                    ' caso seja uma coleção
                    Dim Controles As Object = Nothing
                    Dim Achou As Boolean = False
                    For Each Tipo As Object In TiposG
                        If TypeOf Tipo Is String Then
                            If Raiz.GetType.ToString = Tipo Then
                                Achou = True
                                Exit For
                            End If
                        Else
                            If Raiz.GetType Is Tipo Then
                                Achou = True
                                Exit For
                            End If
                        End If
                    Next
                    If Achou Then
                        Lista.Add(Raiz)
                        If PararNoPrimeiro Then
                            Return Lista
                        End If

                    End If

                    ' caso tenha subcontroles
                    Dim SubControls As Integer = 0
                    Try
                        SubControls = Raiz.controls.count
                    Catch
                    End Try
                    If SubControls > 0 Then
                        For Each Ctl As Object In Raiz.Controls
                            If Not TemNaLista(JaVerificado, Ctl.UNIQUEID) Then
                                CopiaItens(Lista, BuscaTipo(Ctl, PararNoPrimeiro, JaVerificado, Tipos))
                                If Lista.Count > 0 And PararNoPrimeiro Then
                                    Return Lista
                                End If
                            End If
                        Next
                    End If
                End If
                Return Lista
            End Function




            ''' <summary>
            ''' Obtém lista de controlcollections existentes entre as dependências da Raíz.
            ''' </summary>
            ''' <param name="Raiz">Objeto a partir da qual a busca se iniciará.</param>
            ''' <returns>Retorna uma colleção de Controles ou ControlCollections existentes a partir da raíz sendo a busca recursiva.</returns>
            ''' <remarks></remarks>
            Shared Function Containers(ByVal Raiz As Object) As Object
                Dim Lista As New ArrayList

                ' caso seja uma coleção
                Dim Controles As Object = Nothing
                If TypeOf Raiz Is Web.UI.ControlCollection Or TypeOf Raiz Is Windows.Forms.Form.ControlCollection Then
                    If Raiz.Count > 0 Then
                        Lista.Add(Raiz)

                        For Each ctl As Control In Raiz
                            CopiaItens(Lista, Containers(ctl))
                        Next
                    End If
                Else

                    ' caso tenha subcontroles
                    Dim SubControls As Integer = 0
                    Try
                        SubControls = Raiz.controls.count
                    Catch
                    End Try
                    If SubControls > 0 Then
                        Lista.Add(Raiz) 'controlcollection 
                        For Each Ctl As Object In Raiz.Controls
                            CopiaItens(Lista, Containers(Ctl))
                        Next
                    End If
                End If
                Return Lista
            End Function

            ''' <summary>
            ''' Obtenção de arraylist contendo nomes dos campos conforme prefixo.
            ''' </summary>
            ''' <param name="Container">Container onde ocorre a pesquisa.</param>
            ''' <param name="Prefixo">Prefixo pesquisado.</param>
            ''' <returns>Arraylist contendo nomes dos campos.</returns>
            ''' <remarks></remarks>
            Shared Function Campos(ByVal Container As Object, ByVal Prefixo As String) As ArrayList
                Dim Lista As ArrayList = New ArrayList
                For Each Ctl As Control In Controles(Container, Prefixo)
                    Lista.Add(Prop(Ctl, "ID").Substring(Len(Prefixo)))
                Next
                Return Lista
            End Function
        End Class

        ''' <summary>
        ''' Classe para facilitar acesso aos controles.
        ''' </summary>
        ''' <remarks></remarks>
        Public Class Controle

            ''' <summary>
            ''' Avalia se controle possui conteúdo vazio/nulo/nothing ou similar.
            ''' </summary>
            ''' <param name="Ctl">Controle pesquisado.</param>
            ''' <value>TRUE se conteúdo vazio ou FALSE caso não.</value>
            ''' <returns>TRUE se conteúdo vazio ou FALSE caso não.</returns>
            ''' <remarks></remarks>
            Shared ReadOnly Property EraNulo(ByVal Ctl As Object) As Boolean
                Get
                    Return IsDBNull(ValorAnterior(Ctl))
                End Get
            End Property

            ''' <summary>
            ''' Acesso ao valor anterior daquele controle, registrado como atributo.
            ''' </summary>
            ''' <param name="Ctl">Controle pesquisado.</param>
            ''' <value>Valor anterior a ser definido.</value>
            ''' <returns>Valor anterior obtido.</returns>
            ''' <remarks></remarks>
            Shared Property ValorAnterior(ByVal Ctl As Object) As Object
                Get
                    Return Prop(Ctl, "ValorAnterior")
                End Get
                Set(ByVal value As Object)
                    Prop(Ctl, "ValorAnterior") = CampoParaControle(value, Ctl)
                End Set
            End Property

            ''' <summary>
            ''' Pesquisa de valor atual do controle.
            ''' </summary>
            ''' <param name="ctl">Controle pesquisado.</param>
            ''' <param name="RegNovo">TRUE caso seja registro novo (conteúdo anterior inexistente, é claro!).</param>
            ''' <value>Valor atual do controle pesquisado.</value>
            ''' <returns>Valor atual do controle pesquisado.</returns>
            ''' <remarks></remarks>
            Shared Property ValorAtual(ByVal ctl As Object, Optional ByVal RegNovo As Boolean = False) As Object
                Get
                    If TypeOf (ctl) Is DropDownList AndAlso CType(ctl, DropDownList).Text = ComboNull Then
                        Return Convert.DBNull
                    ElseIf Compare(Controle.Auto(ctl), "AUTONUM") And NZV(Prop(ctl, "TEXT"), "[auto]") = "[auto]" Then
                        Return Convert.DBNull
                    ElseIf Compare(Controle.Auto(ctl), "PROXSEQ") And NZV(Prop(ctl, "TEXT"), "[auto]") = "[auto]" Then
                        Return DSProxSeq(Prop(ctl, "campo"), Prop(ctl, "tabela"), Prop(ctl, "strconn"), Nothing)
                    ElseIf Compare(Controle.Formato(ctl), "SENHA") Then
                        Return Icraft.IcftBase.EncrypB(Prop(ctl, "TEXT"))
                    End If
                    Return Valor(ctl, Prop(ctl))
                End Get
                Set(ByVal value As Object)
                    Dim Conteudo As Object = CampoParaControle(value, ctl)

                    ' combobox, tratamento especial: ComboNull, inclui caso não exista e atualiza dependências
                    If TypeOf (ctl) Is DropDownList Then
                        ' se for nulo ou vazio, considera --
                        If NZ(value, "") = "" Then
                            value = ComboNull
                        End If
                        Dim Lista As DropDownList = ctl

                        ' atualiza
                        Prop(Lista) = Conteudo

                        ' caso não tenha na lista, inclui
                        If Prop(Lista) <> Conteudo Then
                            Lista.Items.Add(value)
                            Prop(Lista) = Conteudo
                        End If

                        ' verifica se existem dependências
                        AtualizouControle(Lista)
                        Exit Property
                    End If

                    ' para demais controles, atualiza
                    If Controle.Auto(ctl) <> "" And NZ(Conteudo, "") = "" Then
                        Prop(ctl) = "[auto]"
                    ElseIf RegNovo And NZ(Conteudo, "") = "" Then
                        If Compare(Controle.ValorPadrao(ctl), "[:NOW]") Then
                            Prop(ctl) = Format(Now, Controle.Formato(ctl))
                        ElseIf Compare(Controle.ValorPadrao(ctl), "[:IP]") Then
                            Prop(ctl) = ctl.Page.Request.UserHostAddress
                            Try
                                Prop(ctl) &= " (" & Logon(ctl.Page).Usuario & ")"
                            Catch
                            End Try
                        Else
                            Prop(ctl) = Controle.ValorPadrao(ctl)
                        End If
                    Else
                        Prop(ctl) = Conteudo
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Conteúdo atual do controle.
            ''' </summary>
            ''' <param name="Ctl">Controle pesquisado.</param>
            ''' <param name="Texto">Conteúdo caso campo esteja vazio (valor default).</param>
            ''' <value>Valor atual do campo lido ou para definição.</value>
            ''' <returns>Valor atual do campo lido ou para definição.</returns>
            ''' <remarks></remarks>
            Shared ReadOnly Property Valor(ByVal Ctl As Object, ByVal Texto As Object) As Object
                Get
                    If NZ(Texto, "") = "" Then
                        Return Convert.DBNull
                    End If
                    Dim Tipo As String = NZ(Controle.Tipo(Ctl), "System.String")
                    If Compare(Tipo, "System.Byte[]") Then
                        Return ObjectToByteArray(Texto)
                    ElseIf Compare(Tipo, "System.DateTime") Then

                        ' planejamento de formato das datas
                        ' cada formato deve ser previsto para garantia de retorno correto
                        Dim m As Match = RegexGroup(Texto, "(?<dia>\d{1,2})/(?<mes>\d{1,2})/(?<ano>\d{2,4})(?<compl>$|.*)")
                        If m.Captures.Count > 0 Then
                            Return CType(m.Groups("ano").Value & "-" & m.Groups("mes").Value & "-" & m.Groups("dia").Value & " " & m.Groups("compl").Value, DateTime)
                        End If
                        Return CType(Texto, DateTime)
                    ElseIf TypeOf (Ctl) Is DropDownList AndAlso Texto = ComboNull Then
                        Return Convert.DBNull
                    End If
                    Return CTypeStr(Texto, Tipo)
                End Get
            End Property

            ''' <summary>
            ''' Prepara conteúdo a ser definido em campo (não define, apenas prepara).
            ''' </summary>
            ''' <param name="Valor">Valor a ser atribuído ao controle.</param>
            ''' <param name="Ctl">Controle que receberá a definição.</param>
            ''' <returns>Conteúdo definido para o controle.</returns>
            ''' <remarks></remarks>
            Shared Function CampoParaControle(ByVal Valor As Object, ByVal Ctl As Object) As String
                If NZ(Valor, "") = "" Then
                    Return ""
                End If
                Dim Tipo As String = Controle.Tipo(Ctl)
                If Compare(Tipo, "System.DateTime") Then
                    Return Format(Valor, Controle.Formato(Ctl))
                ElseIf Compare(Tipo, "System.Byte[]") Then
                    Return NZ(Valor, "")
                ElseIf Compare(Tipo, "System.Boolean") Then
                    Valor = NZ(Valor, "")
                    Return CType(Valor, Boolean)
                End If
                Return CType(Valor, String)
            End Function

            ''' <summary>
            ''' Definição de valor padrão para controle.
            ''' </summary>
            ''' <param name="ctl">Controle tratado para valor padrão.</param>
            ''' <value>Valor padrão do controle.</value>
            ''' <returns>Valor padrão do controle.</returns>
            ''' <remarks></remarks>
            Shared Property ValorPadrao(ByVal ctl As Object) As Object
                Get
                    Return Prop(ctl, "ValorPadrao")
                End Get
                Set(ByVal value As Object)
                    Prop(ctl, "ValorPadrao") = value
                End Set
            End Property

            ''' <summary>
            ''' Definição de tipo do controle.
            ''' </summary>
            ''' <param name="Ctl">Controle.</param>
            ''' <value>Tipo do controle.</value>
            ''' <returns>Tipo do controle.</returns>
            ''' <remarks></remarks>
            Shared Property Tipo(ByVal Ctl As Object) As Object
                Get
                    Return Prop(Ctl, "tipo")
                End Get
                Set(ByVal value As Object)
                    Prop(Ctl, "tipo") = value
                End Set
            End Property

            ''' <summary>
            ''' Definição automática para controle.
            ''' </summary>
            ''' <param name="ctl">Controle.</param>
            ''' <value>Definição de forma automática a ser considerada.</value>
            ''' <returns>Definição de forma automática a ser considerada.</returns>
            ''' <remarks></remarks>
            Shared Property Auto(ByVal ctl As Object) As Object
                Get
                    Return Prop(ctl, "Auto")
                End Get
                Set(ByVal value As Object)
                    Prop(ctl, "Auto") = value
                End Set
            End Property

            ''' <summary>
            ''' Formato do controle.
            ''' </summary>
            ''' <param name="Ctl">Controle tratado.</param>
            ''' <value>Formato registrado para aquele controle.</value>
            ''' <returns>Formato registrado para aquele controle.</returns>
            ''' <remarks></remarks>
            Shared Property Formato(ByVal Ctl As Object) As Object
                ' não esquecer de prever formato em formato, mascara
                ' ver no javascript validaentrada também
                Get
                    Dim Forma As String = Prop(Ctl, "Formato")
                    If Forma = "" Then
                        Dim Tipo As String = Controle.Tipo(Ctl)
                        If Compare(Tipo, "System.Int32") Or Compare(Tipo, "System.Int16") Then
                            Forma = "INTEIRO"
                        ElseIf Compare(Tipo, "System.DateTime") Then
                            Forma = "dd\/MM\/yyyy"
                        ElseIf Compare(Tipo, "System.Double") Or Compare(Tipo, "System.Single") Then
                            Forma = "REAL"
                        End If
                    End If
                    Return Forma
                End Get
                Set(ByVal value As Object)
                    Prop(Ctl, "Formato") = value
                End Set
            End Property

            ''' <summary>
            ''' Definição de máscar de preenchimento para controle.
            ''' </summary>
            ''' <param name="Ctl">Controle tratado.</param>
            ''' <value>Conteúdo da máscara do controle.</value>
            ''' <returns>Conteúdo da máscara do controle.</returns>
            ''' <remarks></remarks>
            Shared ReadOnly Property Mascara(ByVal Ctl As Object) As String
                ' não esquecer de prever formato em formato, mascara e mascaraprogress
                ' ver no javascript validaentrada também
                Get
                    Dim Forma As String = Prop(Ctl, "Formato")

                    If Forma = "" Then
                        Dim Tipo As String = Controle.Tipo(Ctl)
                        If Compare(Tipo, "System.Int32") Or Compare(Tipo, "System.Int16") Then
                            Forma = "0"
                        ElseIf Compare(Tipo, "System.DateTime") Then
                            Forma = "dd\/MM\/yyyy"
                        ElseIf Compare(Tipo, "System.Double") Or Compare(Tipo, "System.Single") Then
                            Forma = "0.#########"
                        End If
                    ElseIf Forma = "INTEIRO" Then
                        Forma = "0"
                    ElseIf Forma = "REAL" Then
                        Forma = "0.#########"
                    ElseIf Compare(Forma, "HTML") Or Compare(Forma, "MEMO") Then
                        Forma = ""
                    End If

                    Return Forma
                End Get
            End Property

            ''' <summary>
            ''' Definição de máscara para entrada progressiva, que formata campo e valida mediante entrada de caracter.
            ''' </summary>
            ''' <param name="Ctl">Controle tratado.</param>
            ''' <value>Valor para máscara de entrada.</value>
            ''' <returns>Valor da máscara de entrada.</returns>
            ''' <remarks></remarks>
            Shared ReadOnly Property MascaraProgress(ByVal Ctl As Object) As String
                ' não esquecer de prever formato em formato, mascara
                ' mascaramento progressivo deve ser obrigatoriamente previsto
                ' ver no javascript validaentrada também
                Get
                    Dim Forma As String = Mascara(Ctl)

                    ' tratamento especial do campo inteiro
                    Dim Espec As String = ""
                    If Microsoft.VisualBasic.Left(Forma, 1) = ">" Then ' tudo maiúsculo
                        Espec = Microsoft.VisualBasic.Left(Forma, 1)
                        Forma = Mid(Forma, 2)
                    End If

                    If Forma = "0" Or Forma = "INTEGER" Then
                        Return Espec & "[-+]{0,1}[0-9]*"
                    ElseIf Forma = "dd\/MM\/yyyy" Then
                        Return Espec & "[0-9]{1,2}($|/($|[0-9]{1,2}($|/($|[0-9]{0,4}))))"
                    ElseIf Forma = "0.#########" Or Forma = "REAL" Then
                        Return Espec & "[-+]{0,1}[0-9]*[\\.,]{0,1}[0-9]*"
                    ElseIf Forma = "CURRENCY" Then
                        Return Espec & "[-+]{0,1}[0-9]*($|[\\.,]{0,1}($|[0-9]{1,2}))"
                    ElseIf Forma = "MM\/yyyy" Then
                        Return Espec & "[0-9]{1,2}($|/($|[0-9]{0,4}))"
                    ElseIf Forma = "dd\/MM\/yyyy HH:mm:ss" Then
                        Return Espec & "[0-9]{1,2}($|/($|[0-9]{1,2}($|/($|[0-9]{0,4}($| ($|[0-9]{1,2}($|:($|[0-9]{1,2}($|:($|[0-9]{1,2}))))))))))"
                    ElseIf Forma = "dd\/MM\/yyyy HH:mm" Then
                        Return Espec & "[0-9]{1,2}($|/($|[0-9]{1,2}($|/($|[0-9]{0,4}($| ($|[0-9]{1,2}($|:($|[0-9]{1,2}))))))))"
                    End If
                    Return Espec
                End Get
            End Property

            ''' <summary>
            ''' Retorna texto javascript para atribuição da máscara, considerando funções em icraft.js.
            ''' </summary>
            ''' <param name="Ctl">Controle tratado.</param>
            ''' <remarks></remarks>
            Shared Sub AplicaMascara(ByVal Ctl As Object)
                Prop(Ctl, "OnKeyPress") = "return EntraMasc(this,'" & Controle.MascaraProgress(Ctl).Replace("\", "\\") & "',event)"
                Prop(Ctl, "OnBlur") = "return ValidaMasc(this,'" & Controle.Formato(Ctl).ToString.ToLower & "',event)"

                Dim Expr As String = Controle.Mascara(Ctl)
                Expr = Replace(Expr, "\", "")
                Expr = Replace(Expr, ".", ",")
                Expr = Replace(Expr, ">", "letras maiúsculas")
                Expr = Replace(Expr, ";CAMINHO:", " em ")
                Expr = Replace(Expr, ";SALVASEMCAMINHO:TRUE", "")
                Expr = Replace(Expr, ";MASCARA:", "/")
                Prop(Ctl, "ToolTip") = Prop(Ctl, "ToolTip") & " [" & Expr & "]"
                IncluiScript(Ctl.Page, "Icraft.js")
            End Sub
        End Class

        ''' <summary>
        ''' Classe que armazena detalhes sobre elemento de estilo do tipo "height:300px".
        ''' </summary>
        ''' <remarks></remarks>
        Class ElementoStr
            Private _nome As String = ""
            Private _conteudo As String = ""
            Private _separador As String
            Private _gex_valor_unid As String = "([-0-9.]+)(px|PX)?"
            Private _opera As ElementoStrOpera = ElementoStrOpera.Atribui

            ''' <summary>
            ''' Retorna string representando a forma de estilo.
            ''' </summary>
            ''' <returns>String representando a forma de estilo.</returns>
            ''' <remarks></remarks>
            Overrides Function ToString() As String
                Return _nome & _separador & _conteudo
            End Function

            ''' <summary>
            ''' Cria nova forma de estilo.
            ''' </summary>
            ''' <param name="AtributoStr">Atributo.</param>
            ''' <param name="Separador">Separador.</param>
            ''' <remarks></remarks>
            Sub New(ByVal AtributoStr As String, Optional ByVal Separador As String = ":")
                _separador = Separador
                AtributoStr = NZ(AtributoStr, "")
                Dim pos As Integer = AtributoStr.IndexOf(_separador)
                If pos = -1 Then
                    Conteudo = AtributoStr.Trim
                Else
                    Nome = AtributoStr.Substring(0, pos).Trim
                    Conteudo = AtributoStr.Substring(pos + 1).Trim
                End If
            End Sub

            ''' <summary>
            ''' Nome do atributo, parte à esqueda na definição.
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Property Nome() As String
                Get
                    Return _nome
                End Get
                Set(ByVal value As String)
                    If value.StartsWith("+") Then
                        Operador = ElementoStrOpera.Aumenta
                        _nome = value.Substring(1)
                    ElseIf value.StartsWith("-") Then
                        Operador = ElementoStrOpera.Diminui
                        _nome = value.Substring(1)
                    Else
                        Operador = ElementoStrOpera.Atribui
                        _nome = value
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Conteúdo da definição, parte direita no termo.
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Property Conteudo() As String
                Get
                    Return _conteudo
                End Get
                Set(ByVal value As String)
                    _conteudo = value
                End Set
            End Property

            ''' <summary>
            ''' Extração de valor do conteúdo, quando for acompanhado de termos como "PX" (pixels).
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Property ConteudoValor() As String
                Get
                    Return RegexGroup(Conteudo, _gex_valor_unid, 1).Value
                End Get
                Set(ByVal value As String)
                    Conteudo = RegexGroupReplace(Conteudo, _gex_valor_unid, value, 1)
                End Set
            End Property

            ''' <summary>
            ''' Unidade do conteúdo, como "PX" (pixels).
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Property ConteudoUnidade() As String
                Get
                    Return RegexGroup(Conteudo, _gex_valor_unid, 2).Value
                End Get
                Set(ByVal value As String)
                    Conteudo = RegexGroupReplace(Conteudo, _gex_valor_unid, value, 2)
                End Set
            End Property

            ''' <summary>
            ''' Acesso ao atributo definido para a classe.
            ''' </summary>
            ''' <param name="Nome">Nome do atributo.</param>
            ''' <value>Valor do atributo.</value>
            ''' <returns>Valor do atributo.</returns>
            ''' <remarks></remarks>
            Default Property Attributes(ByVal Nome As String) As String
                Get
                    If Compare(Nome, "Nome") Then
                        Return Me.Nome
                    ElseIf Compare(Nome, "Conteudo") Then
                        Return Conteudo
                    ElseIf Compare(Nome, "ConteudoValor") Then
                        Return ConteudoValor
                    ElseIf Compare(Nome, "ConteudoUnidade") Then
                        Return ConteudoUnidade
                    Else
                        Err.Raise(20000, MyBase.GetType.ToString, "Atributo '" & Nome & "' inválido para objeto " & Me.GetType.ToString & ".")
                    End If
                    Return Nothing
                End Get

                Set(ByVal value As String)
                    If Compare(Nome, "Nome") Then
                        Nome = value
                    ElseIf Compare(Nome, "Conteudo") Then
                        Conteudo = value
                    ElseIf Compare(Nome, "ConteudoValor") Then
                        ConteudoValor = value
                    ElseIf Compare(Nome, "ConteudoUnidade") Then
                        ConteudoUnidade = value
                    Else
                        Err.Raise(20000, MyBase.GetType.ToString, "Atributo " & value & " inválido para objeto " & Me.GetType.ToString & ".")
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Método de operação entre classes, para soma ou exclusão.
            ''' </summary>
            ''' <value>Tipo de operação.</value>
            ''' <returns>Tipo de operação.</returns>
            ''' <remarks></remarks>
            Property Operador() As ElementoStrOpera
                Get
                    Return _opera
                End Get
                Set(ByVal value As ElementoStrOpera)
                    _opera = value
                End Set
            End Property
        End Class

        ''' <summary>
        ''' Classe para armazenar e operar elementostr.
        ''' </summary>
        ''' <remarks></remarks>
        Class ElementosStr
            Private _atributosstr As List(Of ElementoStr) = New List(Of ElementoStr)
            Private _separador As String
            Private _separadorexpr As String
            Private _itemseparador As String
            Private _itemseparadorexpr As String

            ''' <summary>
            ''' Lista de elementos.
            ''' </summary>
            ''' <value>Lista de elementos.</value>
            ''' <returns>Lista de elementos.</returns>
            ''' <remarks></remarks>
            ReadOnly Property Elementos() As List(Of ElementoStr)
                Get
                    Return _atributosstr
                End Get
            End Property

            ''' <summary>
            ''' Especificação de separador.
            ''' </summary>
            ''' <value>Texto contendo separador.</value>
            ''' <returns>Texto contendo separador.</returns>
            ''' <remarks></remarks>
            Property Separador() As String
                Get
                    Return _separador
                End Get
                Set(ByVal value As String)
                    _separador = value
                    _separadorexpr = SeparaExpr(value)
                End Set
            End Property

            ''' <summary>
            ''' Separador de itens.
            ''' </summary>
            ''' <value>Separador de itens.</value>
            ''' <returns>Separador de itens.</returns>
            ''' <remarks></remarks>
            Property ItemSeparador() As String
                Get
                    Return _itemseparador
                End Get
                Set(ByVal value As String)
                    _itemseparador = value
                    _itemseparadorexpr = SeparaExpr(value)
                End Set
            End Property

            ''' <summary>
            ''' Separa parâmetros.
            ''' </summary>
            ''' <param name="Separador">Separador.</param>
            ''' <returns>String contendo itens separados.</returns>
            ''' <remarks></remarks>
            Private Function SeparaExpr(ByVal Separador As String) As String
                Dim Result As String = ""
                For z As Integer = 1 To Separador.Length
                    Dim Letra As String = Mid(Separador, z, 1)
                    If InStr(".\()^|[]+" + Chr(13) + Chr(10), Letra) <> 0 Then
                        Result &= "\" & Letra
                    Else
                        Result &= Letra
                    End If
                Next
                Return Result & "*(([^" & Result & "']|'((([^'])|\\')*)')+)" & Result & "*"
            End Function

            ''' <summary>
            ''' Criação da lista de parâmetros com base em texto e separador.
            ''' </summary>
            ''' <param name="AtributosStr">Lista de atributos.</param>
            ''' <param name="SeparadorTxt">Separador (ex.: border:1px  ; padding:1px).</param>
            ''' <param name="ItemSeparadorTxt">Separador de atributo (border  :  1px).</param>
            ''' <remarks></remarks>
            Sub New(ByVal AtributosStr As String, Optional ByVal SeparadorTxt As String = ";", Optional ByVal ItemSeparadorTxt As String = ":")
                Separador = SeparadorTxt
                ItemSeparador = ItemSeparadorTxt
                AddStr(AtributosStr)
            End Sub

            ''' <summary>
            ''' Transforma estilo em string.
            ''' </summary>
            ''' <param name="SeparadorTxt">Separador a ser utilizado.</param>
            ''' <param name="ItemSeparadorTxt">Atribuidor a ser utilizado.</param>
            ''' <returns>Texto representando estilo com separador e atribuidor escolhidos.</returns>
            ''' <remarks></remarks>
            Function ToStyleStr(Optional ByVal SeparadorTxt As String = Nothing, Optional ByVal ItemSeparadorTxt As String = Nothing) As String
                Dim result As String = ""
                For Each Item As ElementoStr In _atributosstr
                    If Item.Conteudo <> "" And Item.Nome <> "" Then
                        result &= IIf(result <> "", IIf(Not IsNothing(SeparadorTxt), SeparadorTxt, Separador), "")
                        result &= Item.Nome & IIf(Not IsNothing(ItemSeparadorTxt), ItemSeparadorTxt, ItemSeparador)
                        If Item.Conteudo.StartsWith("'") And Item.Conteudo.EndsWith("'") Then
                            result &= Item.Conteudo.Substring(1, Item.Conteudo.Length - 2)
                        Else
                            result &= Item.Conteudo
                        End If
                    End If
                Next
                Return result
            End Function

            ''' <summary>
            ''' Texto representando estilo considerando atributos de separação definidos previamente.
            ''' </summary>
            ''' <returns>Texto representando estilo.</returns>
            ''' <remarks></remarks>
            Overrides Function ToString() As String
                Dim result As String = ""
                For Each Item As ElementoStr In _atributosstr
                    If Item.Conteudo <> "" Then
                        result &= IIf(result <> "", Separador, "") & Item.ToString
                    End If
                Next
                Return result
            End Function

            ''' <summary>
            ''' Lista de itens do estilo.
            ''' </summary>
            ''' <param name="Indice">Índice numérico para acesso ao item do estilo.</param>
            ''' <value>Valor a ser definido.</value>
            ''' <returns>Valor obtido a partir do item consultado.</returns>
            ''' <remarks></remarks>
            Default Overloads Property Items(ByVal Indice As Integer) As ElementoStr
                Get
                    Try
                        Return _atributosstr(Indice)
                    Catch
                    End Try
                    Return New ElementoStr(Nothing, ItemSeparador)
                End Get
                Set(ByVal value As ElementoStr)
                    If Indice = -1 Then
                        _atributosstr.Add(value)
                    Else
                        If Indice >= _atributosstr.Count Then
                            For z As Integer = 0 To Indice
                                _atributosstr.Add(Nothing)
                            Next
                        End If
                        _atributosstr(Indice) = value
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Pesquisa de itens por nome do termo em estilo.
            ''' </summary>
            ''' <param name="Nome">Termo pesquisado.</param>
            ''' <value>Valor a ser atribuído.</value>
            ''' <returns>Valor obtido a partir do termo.</returns>
            ''' <remarks></remarks>
            Default Overloads Property Items(ByVal Nome As String) As ElementoStr
                Get
                    Dim result As ElementoStr = ArrayFindByAtt(_atributosstr.ToArray, Nome, "Nome")
                    If IsNothing(result) Then
                        Dim Elem As ElementoStr = New ElementoStr("", ItemSeparador)
                        Elem.Nome = Nome
                        _atributosstr.Add(Elem)
                        Return _atributosstr(_atributosstr.IndexOf(Elem))
                    End If
                    Return result
                End Get
                Set(ByVal value As ElementoStr)
                    Dim pos As Integer = ArrayIndexFindByAtt(_atributosstr.ToArray, Nome, "Nome")
                    Dim Elem As ElementoStr
                    If pos = -1 Then
                        Elem = New ElementoStr(value.ToString, ItemSeparador)
                        _atributosstr.Add(Elem)
                    Else
                        Elem = _atributosstr(pos)
                        If value.Operador = ElementoStrOpera.Aumenta Then
                            Elem.ConteudoValor = Val(Elem.ConteudoValor) + Val(value.ConteudoValor)
                        ElseIf value.Operador = ElementoStrOpera.Diminui Then
                            Elem.ConteudoValor = Val(Elem.ConteudoValor) - Val(value.ConteudoValor)
                        Else
                            Elem.Conteudo = value.Conteudo
                        End If
                    End If
                End Set
            End Property

            ''' <summary>
            ''' Quantidade de itens no estilo.
            ''' </summary>
            ''' <value>Quantidade de itens no estilo.</value>
            ''' <returns>Quantidade de itens no estilo.</returns>
            ''' <remarks></remarks>
            ReadOnly Property Count() As Integer
                Get
                    Return _atributosstr.Count
                End Get
            End Property

            ''' <summary>
            ''' Itens a serem adicionados ao estilo.
            ''' </summary>
            ''' <param name="AtributosStr">Itens a serem adicionados no estilo.</param>
            ''' <remarks></remarks>
            Sub AddStr(ByVal AtributosStr As String)
                For Each Item As Match In Regex.Matches(AtributosStr, _separadorexpr, RegexOptions.Multiline)
                    Dim Elem As ElementoStr = New ElementoStr(Item.Groups(1).Value, ItemSeparador)
                    Items(Elem.Nome) = Elem
                Next
            End Sub

            ''' <summary>
            ''' Verifica a existência ou não do termo no estilo.
            ''' </summary>
            ''' <param name="Nome">Termo a ser pesquisado.</param>
            ''' <value>TRUE caso exista ou FALSE caso não seja encontrado.</value>
            ''' <returns>TRUE caso exista ou FALSE caso não seja encontrado.</returns>
            ''' <remarks></remarks>
            ReadOnly Property Exists(ByVal Nome As String) As Boolean
                Get
                    Return ArrayIndexFindByAtt(_atributosstr.ToArray, Nome, "Nome") <> -1
                End Get
            End Property
        End Class

        ''' <summary>
        ''' Classe para contagem de tempo. TempoDecorrido retorna o total de segundos passados desde o último registro.
        ''' </summary>
        ''' <remarks></remarks>
        Class ContaTempo
            Private _UltimoRegistro As Date

            ''' <summary>
            ''' Momento de último registro, que corresponde ao ponto de inicialização da contagem.
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property UltimoRegistro() As Date
                Get
                    Return _UltimoRegistro
                End Get
            End Property

            ''' <summary>
            ''' Registra momento atual como início da contagem.
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub Registra()
                _UltimoRegistro = Now
            End Sub

            ''' <summary>
            ''' Quantidade de segundos passados desde o último registro.
            ''' </summary>
            ''' <param name="RegistraOutro">Obtém segundos e inicia nova contagem.</param>
            ''' <value>Segundos passados desde o último registro.</value>
            ''' <returns>Segundos passados desde o último registro.</returns>
            ''' <remarks></remarks>
            Public ReadOnly Property TempoDecorrido(Optional ByVal RegistraOutro As Boolean = True) As Double
                Get
                    Dim RegAnterior As Date = _UltimoRegistro
                    If RegistraOutro Then
                        Registra()
                    End If
                    Return (_UltimoRegistro.Ticks - RegAnterior.Ticks) / 10000000
                End Get
            End Property

            ''' <summary>
            ''' Cria contagem de tempo, registrado momento inicial como atual.
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub New()
                Registra()
            End Sub
        End Class

        ''' <summary>
        ''' Operação com estilos permitindo inclusão, exclusão e atualização de conteúdo.
        ''' </summary>
        ''' <remarks></remarks>
        Public Class Estilo

            ''' <summary>
            ''' Soma de estilos.
            ''' </summary>
            ''' <param name="Esse">Conjunto inicial de estilo "border:1px;padding:2px".</param>
            ''' <param name="MaisEsse">Adicional de estilos a serem incluídos no estilo inicial no mesmo formato "margin:1px;background-color:#f0f0f0".</param>
            ''' <returns>Retorna estilo resultante considerando a soma dos dois sendo eliminadas as redundâncias.</returns>
            ''' <remarks></remarks>
            Shared Function Soma(ByVal Esse As String, ByVal MaisEsse As String) As String
                Return Opera(Esse, MaisEsse, "+")
            End Function

            ''' <summary>
            ''' Retira um estilo específico do estilo inicial.
            ''' </summary>
            ''' <param name="Esse">Lista de estilos inicial.</param>
            ''' <param name="MenosEsse">Estilo a ser retirado do inicial.</param>
            ''' <returns>Estilo resultante, sem os termos solicitados para retirada.</returns>
            ''' <remarks></remarks>
            Shared Function Subtrai(ByVal Esse As String, ByVal MenosEsse As String) As String
                Return Opera(Esse, MenosEsse, "-")
            End Function

            ''' <summary>
            ''' Obtém valor do estilo, sendo eliminado o termo de unidade, por exemplo.
            ''' </summary>
            ''' <param name="Param">Termo do qual queremos o valor.</param>
            ''' <returns>Valor daquele termo.</returns>
            ''' <remarks></remarks>
            Shared Function ConteudoValor(ByVal Param As String) As String
                Dim ValUnid As MatchCollection = System.Text.RegularExpressions.Regex.Matches(Param, "([-0-9.]+) ?(px|PX)?", RegexOptions.Multiline)
                Try
                    Return ValUnid(0).Groups(1).Value
                Catch
                End Try
                Return ""
            End Function

            ''' <summary>
            ''' Unidade do termo.
            ''' </summary>
            ''' <param name="Param">Termo desejado.</param>
            ''' <returns>Retorna unidade daquele termo.</returns>
            ''' <remarks></remarks>
            Shared Function ConteudoUnidade(ByVal Param As String) As String
                Dim ValUnid As MatchCollection = System.Text.RegularExpressions.Regex.Matches(Param, "([-0-9.]+) ?(px|PX)?", RegexOptions.Multiline)
                Try
                    Return ValUnid(0).Groups(2).Value
                Catch
                End Try
                Return ""
            End Function

            ''' <summary>
            ''' Nome do termo solicitado.
            ''' </summary>
            ''' <param name="Param">Estilo solicitado.</param>
            ''' <returns>Nome da variável daquele estilo.</returns>
            ''' <remarks></remarks>
            Shared Function Variavel(ByVal Param As String) As String
                Dim ParamSep As MatchCollection = System.Text.RegularExpressions.Regex.Matches(Param, "([^:]*):([^;]*)(;|$)", RegexOptions.Multiline)
                Try
                    Return ParamSep(0).Groups(1).Value
                Catch
                End Try
                Return ""
            End Function

            ''' <summary>
            ''' Conteúdo do termo solicitado.
            ''' </summary>
            ''' <param name="Param">Termo a ser tratado.</param>
            ''' <returns>Conteúdo definido para aquele termo.</returns>
            ''' <remarks></remarks>
            Shared Function Conteudo(ByVal Param As String) As String
                Dim ParamSep As MatchCollection = System.Text.RegularExpressions.Regex.Matches(Param, "([^:]*):([^;]*)(;|$)", RegexOptions.Multiline)
                Try
                    Return ParamSep(0).Groups(2).Value
                Catch
                End Try
                Return ""
            End Function


            ''' <summary>
            ''' Opera dois estilos com base em um operador.
            ''' </summary>
            ''' <param name="Esse">Primeiro estilo.</param>
            ''' <param name="ComEsse">Segundo estilo.</param>
            ''' <param name="Operador">Tipo de operação.</param>
            ''' <returns>Estilo resultante.</returns>
            ''' <remarks></remarks>
            Shared Function Opera(ByVal Esse As String, ByVal ComEsse As String, ByVal Operador As String) As String
                Dim ValUnidEsse As MatchCollection = System.Text.RegularExpressions.Regex.Matches(Esse, "([-0-9.]+) ?(px|PX)?", RegexOptions.Multiline)
                Dim ValUnidComEsse As MatchCollection = System.Text.RegularExpressions.Regex.Matches(ComEsse, "([-0-9.]+) ?(px|PX)?", RegexOptions.Multiline)

                If ValUnidEsse.Count = 0 Then
                    Return ComEsse
                ElseIf ValUnidEsse.Count = 0 Then
                    Return Esse
                Else
                    Dim ValResult As Integer = Val(ValUnidEsse(0).Groups(1).Value)
                    If String.Compare(ValUnidEsse(0).Groups(2).Value, "px", True) = 0 And String.Compare(ValUnidComEsse(0).Groups(2).Value, "px", True) = 0 Then
                        If Operador = "+" Then
                            ValResult = Val(ValUnidEsse(0).Groups(1).Value) + Val(ValUnidComEsse(0).Groups(1).Value)
                        ElseIf Operador = "-" Then
                            ValResult = Val(ValUnidEsse(0).Groups(1).Value) - Val(ValUnidComEsse(0).Groups(1).Value)
                        End If
                    End If
                    Return ValResult.ToString & ValUnidEsse(0).Groups(2).Value
                End If
            End Function

            ''' <summary>
            ''' Concatena estilos com base em um array de objetos.
            ''' </summary>
            ''' <param name="PARAM">Lista de estilos a serem concatenados.</param>
            ''' <returns>String consolidando todos os estilos concatenados.</returns>
            ''' <remarks></remarks>
            Shared Function Concatena(ByVal ParamArray PARAM() As Object) As String
                Dim result As New Dictionary(Of String, String)

                Dim estilos As ArrayList = ParamArrayToArrayList(PARAM)
                For Each estilo As String In estilos

                    For Each ParamEstilo As Match In System.Text.RegularExpressions.Regex.Matches(estilo, "([^:]*):([^;]*)(;|$)", RegexOptions.Multiline)
                        Dim conteudo As String = ParamEstilo.Groups(2).Value.Trim
                        Dim variav As String
                        Dim anterior As String
                        If ParamEstilo.Groups(1).Value.StartsWith("+") Then
                            variav = ParamEstilo.Groups(1).Value.Substring(1).Trim
                            If result.ContainsKey(variav) Then
                                anterior = result(variav)
                                result.Remove(variav)
                                result.Add(variav, Soma(anterior, conteudo))
                            Else
                                result.Add(variav, conteudo)
                            End If
                        ElseIf ParamEstilo.Groups(1).Value.StartsWith("-") Then
                            variav = ParamEstilo.Groups(1).Value.Substring(1).Trim
                            If result.ContainsKey(variav) Then
                                anterior = result(variav)
                                result.Remove(variav)
                                result.Add(variav, Subtrai(anterior, conteudo))
                            Else
                                result.Add(variav, conteudo)
                            End If
                        Else
                            variav = ParamEstilo.Groups(1).Value.Trim
                            If result.ContainsKey(variav) Then
                                result.Remove(variav)
                                result.Add(variav, conteudo)
                            Else
                                result.Add(variav, conteudo)
                            End If
                        End If
                    Next
                Next

                Dim retorno As String = ""
                For Each item As String In result.Keys
                    retorno &= IIf(retorno <> "", ";", "") & item & ":" & result.Item(item)
                Next
                Return retorno
            End Function
        End Class

        ' classe e funções para apresentação de páginas, controles etc.
        ''' <summary>
        ''' Classe para tratar sitemap como array de urls retornando anterior e próximo.
        ''' </summary>
        ''' <remarks></remarks>
        Class MapPath
            Private _urls As ArrayList
            Sub New(ByVal DiretorioMap As String)
                Dim sm As DataSet = New DataSet
                sm.ReadXml(DiretorioMap & "web.sitemap")
                _urls = RegexToArrayList(sm.GetXml, "url=""([^""]*)""", 1, "value")
            End Sub
            ReadOnly Property Proximo(ByVal Pag As String, Optional ByVal Circular As Boolean = True) As String
                Get
                    Dim pos As Integer = _urls.IndexOf(Pag)
                    If pos = -1 Then
                        pos = Count - 1
                    Else
                        pos = pos + 1
                        If pos >= Count Then
                            If Circular Then
                                pos = 0
                            Else
                                pos = Count - 1
                            End If
                        End If
                    End If
                    Return _urls(pos)
                End Get
            End Property
            ReadOnly Property Anterior(ByVal Pag As String, Optional ByVal Circular As Boolean = True) As String
                Get
                    Dim pos As Integer = _urls.IndexOf(Pag)
                    If pos = -1 Then
                        pos = 0
                    Else
                        pos = pos - 1
                        If pos < 0 Then
                            If Circular Then
                                pos = Count - 1
                            Else
                                pos = 0
                            End If
                        End If
                    End If
                    Return _urls(pos)
                End Get
            End Property
            ReadOnly Property Count() As Integer
                Get
                    Return _urls.Count
                End Get
            End Property
            ReadOnly Property IndexOf(ByVal Pag As String) As Integer
                Get
                    Return _urls(Pag)
                End Get
            End Property
            ReadOnly Property Expressao(ByVal Pag As String) As String
                Get
                    Dim pos As Integer = _urls.IndexOf(Pag)
                    If pos = -1 Then
                        Return "Total " & _urls.Count
                    End If
                    Return pos + 1 & "/" & _urls.Count
                End Get
            End Property
        End Class

        ''' <summary>
        ''' Array valorado, para pesquisas de um modo geral, onde aplicamos um valor de atributo
        ''' para substituição do termo predicado
        ''' IMPORTANTE:: termo iniciado, mas a conclusão final depende de testes de performance 
        '''  entre a utilização desta com resize e a utilização do PARAMTOARRAYLIST, já com todas
        '''  as propriedades declaradas nativamente.
        ''' </summary>
        ''' <remarks></remarks>
        Class ArrayV
            Private ArrayOrigem() As Object
            Private Conteudo As Object
            Private Function PredExists(ByVal Obj As Object) As Boolean
                Return Obj = Conteudo
            End Function
            Public Function Exists(ByVal Conteudo As Object, Optional ByRef ArrayOrigem() As Object = Nothing) As Boolean
                If Not IsNothing(ArrayOrigem) Then
                    Me.ArrayOrigem = ArrayOrigem
                End If
                Me.Conteudo = Conteudo
                Return Array.Exists(ArrayOrigem, AddressOf PredExists)
            End Function
            Sub New(ByRef ArrayOrigem As Object)
                Me.ArrayOrigem = ArrayOrigem
            End Sub
            ' add resize
            ' ver outras em arraylist...
        End Class

        Class SiteMapProviderGeral
            Inherits SiteMapProvider
            Private _Fonte As New TreeNodeCollection
            Private _HomeDescr As String
            Private _HomeURL As String
            Private _HomeToolTip As String
            Private _RetiraTags As Boolean
            Private Function Urls(Optional ByVal rawUrl As String = "") As String
                Dim Ret As String = rawUrl
                If rawUrl <> HttpContext.Current.Request.AppRelativeCurrentExecutionFilePath Then
                    Ret &= IIf(Ret <> "", ";", "") & HttpContext.Current.Request.AppRelativeCurrentExecutionFilePath
                End If
                If rawUrl <> HttpContext.Current.Request.Url.LocalPath Then
                    Ret &= IIf(Ret <> "", ";", "") & HttpContext.Current.Request.Url.LocalPath
                End If
                Return Ret
            End Function

            ''' <summary>
            ''' Classe SiteMapProviderGeral está preparada para ICFTMENU, TREEEVIEW ou TREENODE.
            ''' </summary>
            ''' <param name="Fonte">ICFTMENU, TREEEVIEW ou TREENODE.</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal Fonte As Object, Optional ByVal IncluiHome As Boolean = True, Optional ByVal HomeDescr As String = "HOME", Optional ByVal HomeURL As String = "~/", Optional ByVal HomeToolTip As String = "Página principal.", Optional ByVal RetiraTags As Boolean = False)
                If IncluiHome = True Then
                    _HomeDescr = HomeDescr
                    _HomeURL = HomeURL
                    _HomeToolTip = HomeToolTip
                    _RetiraTags = RetiraTags
                Else
                    _HomeDescr = ""
                End If
                If Fonte.GetType.ToString = TipoTxtIcftMenu Then
                    _Fonte = Fonte.Arvore.Nodes
                ElseIf TypeOf Fonte Is TreeView Then
                    _Fonte = CType(Fonte, TreeView).Nodes
                ElseIf TypeOf Fonte Is TreeNode Then
                    _Fonte = New TreeNodeCollection(CType(Fonte, TreeNode))
                Else
                    _Fonte = Nothing
                End If
            End Sub

            Private Function CriaNodeMapHome() As SiteMapNode
                If _HomeDescr <> "" Then
                    Return CriaNodeMap(New SiteMapNode(Me, "HOME", _HomeURL, _HomeDescr, _HomeToolTip))
                End If
                Return Nothing
            End Function
            Private Function CriaNodeMapFromTreeNode(ByVal Node As TreeNode, Optional ByVal NodeMapFilhos As Object = Nothing, Optional ByVal InsereFilhos As Boolean = True) As SiteMapNode
                Dim Texto As String = Node.Text
                If _RetiraTags Then
                    Texto = Trim(RegexGroup(Texto, "( )*(<[^>]*>)*( )*([^<]*)", 4).Value)
                End If

                Dim NodeMap As SiteMapNode = CriaNodeMap(New SiteMapNode(Me, Node.ValuePath, Node.NavigateUrl, Texto, Node.ToolTip), NodeMapFilhos, InsereFilhos)
                Return NodeMap
            End Function
            Private Function CriaNodeMap(ByVal NodeMap As SiteMapNode, Optional ByVal NodeMapFilhos As Object = Nothing, Optional ByVal InsereFilhos As Boolean = True) As SiteMapNode
                ' nodemap vem com o tal

                ' isso permite que eu passe um ou uma coleção de sitemapnodes
                If Not IsNothing(NodeMapFilhos) Then
                    If TypeOf NodeMapFilhos Is SiteMapNode Then
                        NodeMapFilhos = New SiteMapNodeCollection(CType(NodeMapFilhos, SiteMapNode))
                    End If
                    NodeMap.ChildNodes = New SiteMapNodeCollection
                    NodeMap.ChildNodes.AddRange(NodeMapFilhos)
                End If

                Return NodeMap
            End Function
            Private Function CriaNodes(ByVal Nodes As TreeNodeCollection) As SiteMapNodeCollection
                Dim NodeMaps As New SiteMapNodeCollection
                For Each Node As TreeNode In Nodes
                    NodeMaps.Add(CriaNodeMapFromTreeNode(Node))
                Next
                Return NodeMaps
            End Function

            ' INTERFACE 
            Public Overloads Overrides Function FindSiteMapNode(ByVal rawUrl As String) As System.Web.SiteMapNode
                If Not IsNothing(_Fonte) Then
                    Dim achou As TreeNode = ProcuraNode(_Fonte, NodeCampo.NavigateUrl, Urls(rawUrl))
                    If Not IsNothing(achou) Then
                        Return CriaNodeMapFromTreeNode(achou)
                    End If
                End If
                Return Nothing
            End Function
            Public Overrides Function GetChildNodes(ByVal node As System.Web.SiteMapNode) As System.Web.SiteMapNodeCollection
                If Not IsNothing(_Fonte) Then
                    If node.Key = "HOME" Then
                        Return CriaNodes(_Fonte)
                    Else
                        Dim achou As TreeNode = ProcuraNode(_Fonte, NodeCampo.ValuePath, node.Key)
                        If Not IsNothing(achou) Then
                            Return CriaNodes(achou.ChildNodes)
                        End If
                    End If
                End If
                Return Nothing
            End Function
            Public Overrides Function GetParentNode(ByVal node As System.Web.SiteMapNode) As System.Web.SiteMapNode
                If Not IsNothing(_Fonte) AndAlso Not node.Key = "HOME" Then
                    Dim achou As TreeNode = ProcuraNode(_Fonte, NodeCampo.ValuePath, node.Key)
                    If Not IsNothing(achou) AndAlso Not IsNothing(achou.Parent) Then
                        Return CriaNodeMapFromTreeNode(achou.Parent)
                    Else
                        Return CriaNodeMapHome()
                    End If
                End If
                Return Nothing
            End Function
            Protected Overrides Function GetRootNodeCore() As System.Web.SiteMapNode
                If Not IsNothing(_Fonte) Then
                    Return CriaNodeMapHome()
                End If
                Return Nothing
            End Function
        End Class

        ''' <summary>
        ''' Recurso para simplificar a criação de algumas tags html.
        ''' </summary>
        ''' <remarks></remarks>
        Class HTML
            ''' <summary>
            ''' Orientação corresponde ao sentido inicial de preenchimento da tabela. Horizontal para preencher da esquerda para direita e vertical para preencher de cima para baixo.
            ''' </summary>
            ''' <remarks></remarks>
            Enum Table_Sentido
                Horizontal
                Vertical
            End Enum

            ''' <summary>
            ''' a href cria uma referência de link em um determinado código html.
            ''' </summary>
            ''' <param name="html">Código html que aparecerá sublinhado sendo o link para a execução desejada.</param>
            ''' <param name="link">URL ou código de execução desejada.</param>
            ''' <param name="ListaEstilos">Lista de estilos esperada podendo ser ":link", "border:1px solid red;background-color:#ff0000".</param>
            ''' <returns>Um html contendo o código com a referência.</returns>
            ''' <remarks></remarks>
            Shared Function A_HRef(ByVal html As String, ByVal link As String, ByVal ParamArray ListaEstilos() As Object) As String
                Dim Estilos As ArrayList = ParamArrayToArrayList(ListaEstilos)

                Dim Result As String = html
                If link <> "" Then
                    Result = "<a href=" & SqlExpr(link) & StrStyle("link", Estilos) & ">" & HttpUtility.HtmlEncode(html) & "</a>"
                End If
                Return Result
            End Function

            ''' <summary>
            ''' Função privada para busca de estilo.
            ''' </summary>
            ''' <param name="Tipo">Tipo a ser pesquisado: row, col, cell, rowalt, colalt, tab, etc...</param>
            ''' <param name="Estilos">Lista de estilos no formato ":row", "background-color:#f8f8f8", ":rowalt", "background-color:#e8e8e8", ":cell", "padding:4px".</param>
            ''' <returns>Retorna a string contendo a formatação a ser inserida na tag contendo o estilo.</returns>
            ''' <remarks></remarks>
            Private Shared Function StrStyle(ByVal Tipo As String, ByVal Estilos As ArrayList) As String
                Dim Result As String = ""
                If Not IsNothing(Estilos) Then
                    Dim Ini As Integer = 0
                    Do While Ini < Estilos.Count
                        Dim Pos As Integer = Estilos.IndexOf(":" & Tipo, Ini)
                        If Pos >= 0 Then
                            Dim Achou As String = Estilos(Pos + 1)
                            If InStr(Achou, ":") = 0 Then
                                Result &= " class=" & SqlExpr(Achou)
                            ElseIf InStr(Achou, ":") <> 0 Then
                                Result &= " style=" & SqlExpr(Achou)
                            End If
                        Else
                            Pos = Estilos.Count
                        End If
                        Ini = Pos + 2
                    Loop
                End If
                Return Result
            End Function

            ''' <summary>
            ''' Simplifica o processo de obtenção do estilo para linhas alternadas pela inclusão de alt após o tipo.
            ''' </summary>
            ''' <param name="Tipo">Tipo que pode ser row, col, cell, tab etc...</param>
            ''' <param name="Estilos">Lista de estilos onde será procurado o tipo.</param>
            ''' <param name="Nr">Número da linha ou coluna que está sendo criada.</param>
            ''' <returns>Retorna string contendo estilo a ser aplicado no tag.</returns>
            ''' <remarks></remarks>
            Private Shared Function StrStyle(ByVal Tipo As String, ByVal Estilos As ArrayList, ByVal Nr As Integer) As String
                Dim Result As String = ""
                If Not IsNothing(Estilos) Then
                    If (Nr Mod 2) = 0 Then
                        Result = StrStyle(Tipo & "alt", Estilos)
                        If Result <> "" Then
                            Return Result
                        End If
                    End If
                    Return StrStyle(Tipo, Estilos)
                End If
                Return Result
            End Function

            Shared Function TableSemEncode(ByVal Lista As ArrayList, ByVal Sentido As Table_Sentido, ByVal QtdCols As Integer, ByVal QtdLinhas As Integer, ByVal ParamArray ListaEstilos() As Object) As String
                Dim Estilos As ArrayList = ParamArrayToArrayList(ListaEstilos)

                ' estilos compreendem pré-disposições
                '   tab = tabela
                '   col = coluna
                '   colalt = coluna alterana (par)
                '   cell = célula
                '   row = linha
                '   rowalt = linha alternada (par)
                ' formato:
                '              ":row", "background-color:#f8f8f8", ":rowalt", "background-color:#e8e8e8", ":cell", "padding:4px" ...

                ' calcula itens no sentido escolhido
                ' sentido horizontal, prioridade para colunas
                ' sentido vertical, prioridade para linhas

                Dim QtdItensNoSentido As Integer
                If Sentido = Table_Sentido.Horizontal Then
                    QtdItensNoSentido = NZV(QtdCols, Int((Lista.Count - 1) / NZV(QtdLinhas, 1)) + 1)
                Else
                    QtdItensNoSentido = NZV(QtdLinhas, Int((Lista.Count - 1) / NZV(QtdCols, 1)) + 1)
                End If

                Dim Result As String = ""
                Result = "<table" & StrStyle("tab", Estilos) & "> " & vbCrLf

                Dim nl As Integer = 1
                For z As Integer = 0 To IIf(Sentido = Table_Sentido.Horizontal, Lista.Count - 1, QtdItensNoSentido - 1) Step IIf(Sentido = Table_Sentido.Horizontal, QtdItensNoSentido, 1)
                    Result &= "    <tr" & StrStyle("row", Estilos, nl) & ">" & vbCrLf

                    Dim nc As Integer = 1
                    For zz As Integer = 0 To IIf(Sentido = Table_Sentido.Horizontal, QtdItensNoSentido - 1, Lista.Count - 1) Step IIf(Sentido = Table_Sentido.Horizontal, 1, QtdItensNoSentido)
                        Result &= "        <td" & StrStyle("col", Estilos, nc) & StrStyle("cell", Estilos) & ">" & vbCrLf
                        If (z + zz) < Lista.Count Then
                            Result &= NZ(Lista.Item(z + zz), "") & vbCrLf
                        Else
                            Result &= "&nbsp" & vbCrLf
                        End If
                        Result &= "        </td>" & vbCrLf
                        nc += 1
                    Next
                    Result &= "    </tr>" & vbCrLf
                    nl += 1
                Next

                Result &= "</table>"
                Return Result
            End Function

            ''' <summary>
            ''' Cria código de tabela seguindo um determinado sentido considerando uma lista de valores sequenciados.
            ''' </summary>
            ''' <param name="Lista">Lista de conteúdos sequenciados.</param>
            ''' <param name="Sentido">Sentido pode ser horizonta para preenchimento da tabela da esquerda para direita e vertical para preencher de cima para baixo.</param>
            ''' <param name="QtdCols">Quantidade de colunas. Será priorizada esta quantidade caso seja sentido horizontal.</param>
            ''' <param name="QtdLinhas">Quantidade de linhas. Será priorizada esta quantidade caso seja sentido vertical.</param>
            ''' <param name="ListaEstilos">Lista de estilos configurados para preenchimento dos tags table, tr e td.</param>
            ''' <returns>Retorna o código html da tabela contendo os valores.</returns>
            ''' <remarks></remarks>
            Shared Function Table(ByVal Lista As ArrayList, ByVal Sentido As Table_Sentido, ByVal QtdCols As Integer, ByVal QtdLinhas As Integer, ByVal ParamArray ListaEstilos() As Object) As String
                Dim Estilos As ArrayList = ParamArrayToArrayList(ListaEstilos)

                ' estilos compreendem pré-disposições
                '   tab = tabela
                '   col = coluna
                '   colalt = coluna alterana (par)
                '   cell = célula
                '   row = linha
                '   rowalt = linha alternada (par)
                ' para definir estilo, é só mencionar pré-disposições seguindo de ":" mais os itens conforme cláusula style
                ' exemplo: col:background-color:#F0F0F0;border:1px solid red

                ' calcula itens no sentido escolhido
                ' sentido horizontal, prioridade para colunas
                ' sentido vertical, prioridade para linhas

                For z As Integer = 0 To Lista.Count - 1
                    Lista.Item(z) = HttpUtility.HtmlEncode(Lista.Item(z))
                Next

                Return TableSemEncode(Lista, Sentido, QtdCols, QtdLinhas, ListaEstilos)
            End Function

            Shared Function Protege(ByVal Texto As String) As String
                Dim HTTP As HttpContext = HttpContext.Current
                Dim Ret As String = ""

                ' preparação geral
                Ret = HTTP.Server.HtmlEncode(Texto)
                Ret = "<p>" & Ret.Replace(vbCrLf, "</p><p>") & "</p>"

                ' html permitido
                ' vou ativar quando necessário protegidotag, para permitir algumas cláusulas em tag
                Ret = Ret.Replace("&lt;b&gt;", "<strong>")
                Ret = Ret.Replace("&lt;/b&gt;", "</strong>")
                Ret = Ret.Replace("&lt;h1&gt;", "<h1>")
                Ret = Ret.Replace("&lt;/h1&gt;", "</h1>")
                Return Ret
            End Function

            'Shared Function ProtegidoTag(ByVal Texto As String, ByVal RegexTag As String) As String
            '    Dim Result As String = Texto
            '    Dim Menos As Integer = 0
            '    For Each m As Match In System.Text.RegularExpressions.Regex.Matches(Result, Mascara)
            '        Result = Result.Substring(0, m.Groups(Grupo).Index - Menos) & TrocarPara & Result.Substring(m.Groups(Grupo).Index + m.Groups(Grupo).Length - Menos)
            '        Menos += m.Groups(Grupo).Length
            '    Next
            '    Return Result
            'End Function
        End Class



        ''' <summary>
        ''' Substitui um grupo regex no texto.
        ''' </summary>
        ''' <param name="Texto">Texto que será avaliado.</param>
        ''' <param name="Mascara">Expressão regex para grupos.</param>
        ''' <param name="TrocarPara">Texto que substituirá o grupo regex encontrado.</param>
        ''' <param name="Grupo">Grupo regex a ser substituído.</param>
        ''' <returns>Retorna texto com grupo substituído por trocarpara com base na expressão regex.</returns>
        ''' <remarks></remarks>
        Shared Function RegexGroupReplace(ByVal Texto As String, ByVal Mascara As String, ByVal TrocarPara As String, Optional ByVal Grupo As Object = 0) As String
            Dim Result As String = Texto
            Dim Menos As Integer = 0
            For Each m As Match In System.Text.RegularExpressions.Regex.Matches(Result, Mascara)
                Result = Result.Substring(0, m.Groups(Grupo).Index - Menos) & TrocarPara & Result.Substring(m.Groups(Grupo).Index + m.Groups(Grupo).Length - Menos)
                Menos += m.Groups(Grupo).Length
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Retorna um grupo regex.
        ''' </summary>
        ''' <param name="Texto">Texto de onde será extraído o grupo.</param>
        ''' <param name="Mascara">Expressão regex para análise do texto.</param>
        ''' <param name="Grupo">Grupo regex a ser extraído.</param>
        ''' <returns>Retorna grupo regex conforme análise do texto com base na máscara.</returns>
        ''' <remarks></remarks>
        Shared Function RegexGroup(ByVal Texto As String, ByVal Mascara As String, Optional ByVal Grupo As Object = 0) As System.Text.RegularExpressions.Group
            Return System.Text.RegularExpressions.Regex.Match(NZ(Texto, ""), Mascara).Groups(Grupo)
        End Function

        ''' <summary>
        ''' Retorna MATCHES de uma consulta regex (apenas para simplificar código).
        ''' </summary>
        ''' <param name="Texto">Texto a ser pesquisado.</param>
        ''' <param name="Mascara">Máscara utilizada para pesquisa.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Shared Function RegexMatches(ByVal Texto As String, ByVal Mascara As String) As Match
            Return System.Text.RegularExpressions.Regex.Match(NZ(Texto, ""), Mascara)
        End Function

        ''' <summary>
        ''' Interpreta uma string utiliza regex retornando um arraylist mediante a escolha de um grupo.
        ''' </summary>
        ''' <param name="Texto">Texto a ser interpretado.</param>
        ''' <param name="Mascara">Máscara regex utilizada para interpretar o texto.</param>
        ''' <param name="Grupo">Grupo desejado sendo grupo zero o padrão (todo o texto).</param>
        ''' <returns>Retorna um arraylist contendo como índice cada ocorrência no formato da máscara no texto.</returns>
        ''' <remarks></remarks>
        Shared Function RegexToArrayList(ByVal Texto As String, ByVal Mascara As String, Optional ByVal Grupo As Integer = 0, Optional ByVal prop As String = "") As ArrayList
            Dim conjunto As ArrayList = New ArrayList
            For Each Item As Match In System.Text.RegularExpressions.Regex.Matches(Texto, Mascara, RegexOptions.Multiline)
                If prop = "" Then
                    conjunto.Add(Item.Groups(Grupo))
                ElseIf Compare(prop, "value") Then
                    conjunto.Add(Item.Groups(Grupo).Value)
                Else
                    Throw New Exception("Prop " & prop & " inválida para grupo em regextoarraylist.")
                End If
            Next
            Return conjunto
        End Function

        ''' <summary>
        ''' Rotina recursiva para regex.match de tags. Apenas retorna a máscara. Informe níveis possíveis para aquela tag.
        ''' </summary>
        ''' <param name="Tag">Tag inicial que também é utilizada como texto recursivo.</param>
        ''' <param name="Niveis">Níveis possíveis para primeira tag.</param>
        ''' <param name="NivelAtual">Indicador de nível para chave recursiva.</param>
        ''' <returns>Texto contendo máscara.</returns>
        ''' <remarks></remarks>
        Shared Function RegexMascTags(Optional ByVal Tag As String = "", Optional ByVal Niveis As Integer = 1, Optional ByVal NivelAtual As Integer = 1) As String
            Dim Texto As String = ""
            If NivelAtual > Niveis Then
                Return ""
            ElseIf NivelAtual > 1 And NivelAtual <= Niveis Then
                Texto = RegexMascTags(Tag, Niveis, NivelAtual + 1)
                Return "<\<tag>( .*?)?>" & IIf(Texto <> "", "(" & Texto & "|.)", ".") & "*?</\<tag>>"
            End If
            Texto = RegexMascTags(Tag, Niveis, NivelAtual + 1)
            Return "(?is)<(?<tag>" & IIf(Tag <> "", Tag, "[^ >]*") & ")( .*?)?>(?<inner>" & IIf(Texto <> "", "(" & Texto & "|.)", ".") & "*?)</\<tag>>"
        End Function


        ''' <summary>
        ''' Objecto para manipulação de texto como html.
        ''' </summary>
        ''' <remarks></remarks>
        Class RegexHtml
            Enum TipoStatus
                OK
                ErroProx
                ErroDentro
                ErroItem
            End Enum
            Dim _m As Match = Nothing
            Dim Trecho As String = ""
            Dim Niveis As Integer
            Dim Tag As String
            Dim Status As TipoStatus = TipoStatus.OK

            Public Shared Function Obtem(ByVal Texto As String, Optional ByVal Tag As String = "", Optional ByVal Niveis As Integer = 4) As RegexHtml
                Dim r As New RegexHtml(Texto, Tag, Niveis)
                Return r
            End Function
            Sub New(ByVal Texto As String, Optional ByVal Tag As String = "", Optional ByVal Niveis As Integer = 4)
                Me.Trecho = Texto
                Me.Niveis = Niveis
                Me.Tag = Tag
                Status = TipoStatus.OK
            End Sub
            Property M() As Match
                Get
                    Try
                        If Status = TipoStatus.OK Then
                            If IsNothing(_m) Then
                                _m = Regex.Match(Trecho, Icraft.IcftBase.RegexMascTags(Tag, Me.Niveis))
                            End If
                        End If
                        Return _m
                    Catch
                    End Try
                    Return Nothing
                End Get
                Set(ByVal value As Match)
                    _m = value
                End Set
            End Property

            Function Ms() As System.Text.RegularExpressions.MatchCollection
                Return Regex.Matches(Trecho, Icraft.IcftBase.RegexMascTags(Tag, Me.Niveis))
            End Function

            Function Prox() As RegexHtml
                Try
                    If Status = TipoStatus.OK Then
                        M = M.NextMatch()
                    End If
                Catch
                    Status = TipoStatus.ErroProx
                End Try
                Return Me
            End Function
            Function Dentro(Optional ByVal TagDesejada As String = Nothing) As RegexHtml
                Try
                    If Status = TipoStatus.OK Then
                        Trecho = M.Groups("inner").Value
                        If Not IsNothing(TagDesejada) Then
                            Tag = TagDesejada
                        End If
                        M = Regex.Match(Trecho, RegexMascTags(Tag, Niveis))
                    End If
                Catch
                    Status = TipoStatus.ErroDentro
                End Try
                Return Me
            End Function
            Default ReadOnly Property Item(ByVal Indice As Integer) As RegexHtml
                Get
                    Try
                        If Status = TipoStatus.OK Then
                            Dim ms As System.Text.RegularExpressions.MatchCollection = Regex.Matches(Trecho, RegexMascTags(Tag, Me.Niveis))
                            M = ms.Item(Indice)
                        End If
                    Catch
                        Status = TipoStatus.ErroItem
                    End Try
                    Return Me
                End Get
            End Property
            Function Inner() As String
                Try
                    If Status = TipoStatus.OK Then
                        Return M.Groups("inner").Value
                    End If
                Catch
                End Try
                Return Nothing
            End Function
            Function Outer() As String
                Try
                    If Status = TipoStatus.OK Then
                        Return M.Value
                    End If
                Catch
                End Try
                Return Nothing
            End Function
        End Class



        Shared Function PaginacaoComando(ByRef DataRep As Repeater, ByVal Fonte As Object, ByVal NumLinhas As Integer, ByVal PagAtual As Integer, ByVal Adicao As Boolean, ByVal Chave As String, ByRef SinalizaItem As Integer, ByVal Comando As String, ByVal ParamArray Paineis() As Object) As PagedDataSource
            Dim DV As DataView
            If TypeOf (Fonte) Is System.Data.DataView Then
                DV = CType(Fonte, DataView)
            Else
                DV = CType(Fonte, DataSet).Tables(0).DefaultView
            End If

            SinalizaItem = -1
            If Chave <> "" Then
                DV.Sort = Replace(Chave, ";", ",")
                If DV.Count > 0 Then
                    If IsDBNull(DV(DV.Count - 1)(0)) Then
                        DV.Delete(DV.Count - 1)
                    End If
                End If

                If Not IsNothing(Comando) Then
                    If Comando.StartsWith("CH=", StringComparison.OrdinalIgnoreCase) Then
                        Dim pos As Integer = DV.Find(Split(Mid(Comando, 4), "|"))
                        If pos <> -1 Then
                            PagAtual = Int(pos / NumLinhas) + 1  'pag um como inicial
                            SinalizaItem = pos Mod NumLinhas     'item zero como primeiro
                        End If
                    End If
                End If
            End If

            Return Paginacao(DataRep, DV, NumLinhas, PagAtual, Adicao, Paineis)
        End Function
        Shared Function Paginacao(ByRef DataRep As Object, ByVal Fonte As Object, ByVal NumLinhas As Integer, ByVal PagAtualouChave As Object, ByVal Adicao As Boolean, ByVal ParamArray Paineis() As Object) As PagedDataSource
            Dim Pag As PagedDataSource = New PagedDataSource

            Dim DV As DataView = Nothing
            If TypeOf (Fonte) Is System.Data.DataRowCollection OrElse IsNothing(Fonte) Then
                Pag.DataSource = Fonte
            ElseIf TypeOf (Fonte) Is System.Data.DataView OrElse TypeOf (Fonte) Is System.Data.DataSet Then
                If TypeOf (Fonte) Is System.Data.DataView Then
                    DV = CType(Fonte, DataView)
                Else
                    DV = CType(Fonte, DataSet).Tables(0).DefaultView
                End If
                If DV.Count > 0 Then
                    If IsDBNull(DV(DV.Count - 1)(0)) Then
                        DV.Delete(DV.Count - 1)
                    End If
                End If
                If Adicao Then
                    DV.AddNew()
                End If
                Pag.DataSource = DV
            ElseIf Fonte.GetType.GetInterfaces.Contains(GetType(Collections.IEnumerable)) Then
                Dim f As System.Collections.IEnumerable = Fonte
                Pag.DataSource = f
            End If

            Pag.AllowPaging = True
            Pag.PageSize = NumLinhas

            ' calcula última página
            Dim UltPag As Integer = Int((Pag.DataSourceCount - Pag.PageSize - IIf(Adicao, 1, 0)) / Pag.PageSize) + 1

            ' busca página atual
            ' ultpag está com num de 1..
            ' PagAtualouChave também.
            If TypeOf (PagAtualouChave) Is ArrayList AndAlso Not IsNothing(DV) Then
                ' busca pela chave
                PagAtualouChave = DV.Table.Rows.IndexOf(DV.Table.Rows.Find(CType(PagAtualouChave, ArrayList).ToArray)) + 1
            Else
                ' caso seja passado número, ocorre posicionamento por registro
                If PagAtualouChave = 0 Then
                    PagAtualouChave = UltPag
                ElseIf PagAtualouChave > Pag.PageCount Then
                    PagAtualouChave = 1
                ElseIf PagAtualouChave <= -1 Then
                    PagAtualouChave = UltPag + 1
                End If
            End If

            ' currentpageindex vai de zero em diante
            Pag.CurrentPageIndex = PagAtualouChave - 1

            ' currentpage < 0 sign um reg somente (vai para último sempre)
            If Pag.CurrentPageIndex < 0 Then
                Pag.CurrentPageIndex = 0
            End If

            ' se currentpageindex for maior que num pags, nothing
            If Pag.CurrentPageIndex >= Pag.PageCount Or Pag.CurrentPageIndex < 0 Then
                DataRep.DataSource = Nothing
            Else
                DataRep.DataSource = Pag
            End If

            ' ativa desativa botões
            Dim PaineisArr As ArrayList = ParamArrayToArrayList(Paineis)
            For Each painel As Object In PaineisArr
                If TypeOf (painel) Is Panel Then
                    CType(painel, Panel).Enabled = True
                    For Each ctl As Control In CType(painel, Panel).Controls
                        If InStr(ctl.ID, "Anterior") <> 0 Then
                            Prop(ctl, "Enabled") = Not Pag.IsFirstPage
                        ElseIf InStr(ctl.ID, "Proximo") <> 0 Then
                            Prop(ctl, "Enabled") = Not Pag.IsLastPage
                        ElseIf InStr(ctl.ID, "Primeiro") <> 0 Then
                            Prop(ctl, "Enabled") = Not Pag.IsFirstPage
                        ElseIf InStr(ctl.ID, "Ultimo") <> 0 Then
                            Prop(ctl, "Enabled") = Not Pag.IsLastPage
                        ElseIf InStr(ctl.ID, "QtdPags") <> 0 Then
                            Prop(ctl, "Text") = UltPag
                        ElseIf InStr(ctl.ID, "PagAtualouChave") <> 0 Then
                            Prop(ctl, "ValorAnterior") = Prop(ctl, "Text")
                            Prop(ctl, "Text") = PagAtualouChave
                        ElseIf InStr(ctl.ID, "Novo") <> 0 Then
                            Prop(ctl, "Enabled") = Adicao And (Pag.CurrentPageIndex < (Pag.PageCount - 1))
                        End If
                    Next
                End If
            Next

            Return Pag
        End Function

        ''' <summary>
        ''' Retorna url anterior à página informada considerando sitemap.
        ''' </summary>
        ''' <param name="DiretorioMap">Diretório onde pode ser encontrado o sitemap.</param>
        ''' <param name="Pag">Página atual.</param>
        ''' <returns>Retorna página anterior caso a atual seja encontrada ou a primeira caso não seja encontrada.</returns>
        ''' <remarks>Para compatibilidade, melhor utilizar mappath.</remarks>
        Shared Function MapPathAntes(ByVal DiretorioMap As String, ByVal Pag As String) As String
            Dim _map As MapPath = New MapPath(DiretorioMap)
            Return _map.Anterior(Pag.ToLower())
        End Function

        ''' <summary>
        ''' Retorna url posterior à página informada considerando sitemap.
        ''' </summary>
        ''' <param name="DiretorioMap">Diretório onde pode ser encontrado o sitemap.</param>
        ''' <param name="Pag">Página atual.</param>
        ''' <returns>Retorna página posterior caso a atual seja encontrada ou a última caso não seja encontrada.</returns>
        ''' <remarks>Para compatibilidade, melhor utilizar mappath.</remarks>
        Shared Function MapPathDepois(ByVal DiretorioMap As String, ByVal Pag As String) As String
            Dim _map As MapPath = New MapPath(DiretorioMap)
            Return _map.Proximo(Pag.ToLower())
        End Function

        ''' <summary>
        ''' Retorna expressão com página atual e total de páginas.
        ''' </summary>
        ''' <param name="DiretorioMap">Diretório onde pode ser encontrado o sitemap.</param>
        ''' <param name="Pag">Página atual.</param>
        ''' <returns>Retorna uma expressão que mostra página atual e quantidade de páginas.</returns>
        ''' <remarks></remarks>
        Shared Function Paginas(ByVal DiretorioMap As String, ByVal Pag As String) As String
            Dim _map As MapPath = New MapPath(DiretorioMap)
            Return _map.Expressao(Pag.ToLower())
        End Function



        Shared Sub DownloadArquivoTam(ByVal Pagina As Page, ByVal Arquivo As String, Optional ByVal NomeDeDownload As String = "", Optional ByVal Deslocamento As Integer = 0, Optional ByVal Tamanho As Integer = 0)
            Arquivo.Replace("/", "\")
            If NomeDeDownload = "" Then
                NomeDeDownload = Arquivo.Substring(Arquivo.LastIndexOf("\") + 1)
            End If
            Pagina.Response.Clear()
            Pagina.Response.ContentType = "application/octet-stream"
            Pagina.Response.AddHeader("Content-disposition", "attachment; filename=" & NomeDeDownload)

            Dim TamArq As Integer = (New IO.FileInfo(Arquivo)).Length
            If Tamanho = 0 OrElse Tamanho > (TamArq - Deslocamento) Then
                Tamanho = TamArq - Deslocamento
            End If

            Dim Arq As New StreamReader(Arquivo)
            Dim stream As Stream = Arq.BaseStream
            Dim buf(Tamanho * 2) As Byte
            stream.Read(buf, Deslocamento, Tamanho)

            Pagina.Response.AddHeader("Content-Length", Tamanho)
            Pagina.Response.AddHeader("Pragma", "no-cache")
            Pagina.Response.Expires = 0

            Pagina.Response.BinaryWrite(buf)
            Pagina.Response.End()
        End Sub


        Shared Sub DownloadArquivo(ByVal Pagina As Page, ByVal Arquivo As String, Optional ByVal NomeDeDownload As String = "")
            Arquivo.Replace("/", "\")
            If NomeDeDownload = "" Then
                NomeDeDownload = Arquivo.Substring(Arquivo.LastIndexOf("\") + 1)
            End If
            Pagina.Response.Clear()
            Pagina.Response.ContentType = "application/octet-stream"
            Pagina.Response.AddHeader("Content-disposition", "attachment; filename=" & NomeDeDownload)

            Dim TamArq As Integer = (New IO.FileInfo(Arquivo)).Length

            '        Pagina.Response.AddHeader("Content-Length", TamArq)
            Pagina.Response.AddHeader("Pragma", "no-cache")
            Pagina.Response.Expires = 0

            Pagina.Response.TransmitFile(Arquivo)
            Pagina.Response.End()
        End Sub

        Shared Sub DownloadConteudo(ByVal Pagina As Page, ByVal Conteudo As String, Optional ByVal NomeDeDownload As String = "")
            Pagina.Response.Clear()
            Pagina.Response.ContentType = "application/octet-stream"
            Pagina.Response.AddHeader("Content-disposition", "attachment; filename=" & NomeDeDownload)
            Pagina.Response.AddHeader("Pragma", "no-cache")
            Pagina.Response.Expires = 0
            Pagina.Response.Write(Conteudo)
            Pagina.Response.Flush()
            Pagina.Response.End()
        End Sub

        Public Shared Sub DownloadXML(ByVal Pagina As Page, ByVal XML As String)
            Pagina.Response.ContentType = "application/octet-stream"
            Pagina.Response.AddHeader("Content-disposition", "attachment; filename=dados.xml")
            If Not XML.StartsWith("<?xml ") Then
                XML = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & XML
            End If
            ' TESTAR TESTAR TESTAR
            ' Pagina.Response.AddHeader("Content-Length", Len(XML))
            Pagina.Response.AddHeader("Pragma", "no-cache")
            Pagina.Response.Expires = 0
            Pagina.Response.Write(XML)
        End Sub


        Shared Sub DownloadConteudo(ByVal Pagina As Page, ByVal Stream As System.IO.BinaryReader, Optional ByVal NomeDeDownload As String = "")
            Pagina.Response.Clear()
            Pagina.Response.ContentType = "application/octet-stream"
            Pagina.Response.AddHeader("Content-disposition", "attachment; filename=" & NomeDeDownload)
            Pagina.Response.AddHeader("Pragma", "no-cache")
            Pagina.Response.Expires = 0
            Do While Stream.BaseStream.Position < Stream.BaseStream.Length
                Pagina.Response.BinaryWrite(Stream.ReadBytes(10000))
            Loop
            Pagina.Response.Flush()
            Pagina.Response.End()
        End Sub


        ''' <summary>
        ''' Acessa parâmetros armazenados sobre o login do usuário em uma determinada tela de autenticação. Utilize LOGIN para logar e LOGOFF para "deslogar".
        ''' </summary>
        ''' <param name="Pagina">Página do requisitante ('me' ou 'page').</param>
        ''' <value>LogonSession, que contém informações de usuário, momento de logon entre outras.</value>
        ''' <returns>LogonSession, que contém informações de usuário, momento de logon entre outras.</returns>
        ''' <remarks></remarks>
        Shared Property Logon(ByVal Pagina As Page) As LogonSession
            Get
                If Not IsNothing(Pagina) Then
                    Dim Area As String = "logon_" & WebConf("site_nome")
                    Dim LL As LogonSession = Pagina.Session(Area)
                    If IsNothing(LL) Then
                        LL = New LogonSession(Pagina, "", "")
                        Logon(Pagina) = LL
                    End If
                    Return LL
                End If
                Return Nothing
            End Get
            Set(ByVal value As LogonSession)
                If Not IsNothing(Pagina) Then
                    Dim Area As String = "logon_" & WebConf("site_nome")
                    If IsNothing(value) Then
                        Dim NaSessao As LogonSession = Pagina.Session(Area)
                        If Not IsNothing(NaSessao) Then
                            Pagina.Session.Remove(Area)
                        End If
                    Else
                        Pagina.Session(Area) = value
                    End If
                End If
            End Set
        End Property


        ''' <summary>
        ''' Efetua login fazendo registro dos dados de usuário.
        ''' </summary>
        ''' <param name="Pagina">Passa a página em questão (me ou page)</param>
        ''' <param name="Usuario">Usuário que está sendo submetido.</param>
        ''' <param name="Senha">Senha do usuário.</param>
        ''' <param name="StrConn">Nome da conexão no webconfig para consulta de usuário.</param>
        ''' <param name="TabelaUsuario">Tabela onde são armazenados os usuários.</param>
        ''' <param name="CampoUsuario">Campo que armazena o nome do usuário.</param>
        ''' <param name="CampoSenha">Campo que armazena a senha.</param>
        ''' <param name="RedirAuto">Redireciona automaticamente para a página.</param>
        ''' <returns>Retorna true caso ocorra o logon sendo LOGON(SESSION) 
        ''' registrada corretamente e false caso ocorra algum erro.</returns>
        ''' <remarks>A função permite busca mediante o acesso a base. Para isso, 
        ''' utilizar como senha na base: [PROVIDER:strconexao], onde conexão corresponde
        ''' à conexão no webconfig.</remarks>
        Public Shared Function Login(ByVal Pagina As Page, ByVal Usuario As String, ByVal Senha As String, ByVal StrConn As Object, Optional ByVal TabelaUsuario As String = "SYS_CONFIG_USUARIO", Optional ByVal CampoUsuario As String = "USUARIO", Optional ByVal CampoSenha As String = "SENHA", Optional ByVal RedirAuto As Boolean = True, Optional ByVal Grupo As String = "") As Boolean
            Dim ConnW As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn)

            If Not IsNothing(Logon(Pagina)) Then
                Logoff(Pagina)
            End If

            ' valida usuário
            Dim SenhaDB As String = NZ(DSValor(CampoSenha, TabelaUsuario, ConnW, CampoUsuario & "=:CAMPOUSUARIO", ":CAMPOUSUARIO", Usuario), "")
            Dim OutroProv As String = RegexGroup(SenhaDB, "\[PROVIDER:([^\]]+)\]", 1).Value
            If OutroProv <> "" Then
                If WebConn(OutroProv).ProviderName = Oracle Then
                    TabelaUsuario = "ALL_TABLES"
                Else
                    Throw New Exception("Login, necessária definição provider para pesquisa.")
                End If

                Try
                    Dim TestaAcesso As Integer = DSValor("COUNT(*)", TabelaUsuario, OutroProv & ";USER:[:USER];PASSWORD:[:PASSWORD]", "", ":USER", Usuario, ":PASSWORD", Senha)
                Catch
                    Return False
                End Try
            Else
                If NZ(Senha, "<<vazio>>") <> NZ(SenhaDB, "<<vazio>>") Then
                    Return False
                End If
            End If

            LoginCertifica(Pagina, Usuario, Senha, RedirAuto, Grupo)
            Return True
        End Function

        Shared Sub LoginCertifica(ByVal Pagina As Page, ByVal Usuario As String, ByVal Senha As String, Optional ByVal RedirAuto As Boolean = True, Optional ByVal Grupo As String = "")

            ' certifica logon
            Dim l As New LogonSession(Pagina, Usuario, Senha)

            Dim TckNome As String

            If Grupo = "" Then
                TckNome = Usuario
            Else
                TckNome = Grupo
            End If

            Dim TckAuth As New FormsAuthenticationTicket(1, TckNome, Now, Now.AddHours(2), False, TckNome)
            Dim TckEncr As String = FormsAuthentication.Encrypt(TckAuth)
            Dim HttpC As HttpCookie = New HttpCookie(FormsAuthentication.FormsCookieName, TckEncr)
            Pagina.Response.Cookies.Add(HttpC)
            Pagina.Session.Timeout = 120

            ' salva na sessão
            Logon(Pagina) = l

            ' redireciona para formulário solicitado
            If RedirAuto Then
                FormsAuthentication.RedirectFromLoginPage(TckNome, False)
            End If
        End Sub

        ''' <summary>
        ''' Cancela sessão existente.
        ''' </summary>
        ''' <param name="Pagina">Objeto da página atual (me ou page).</param>
        ''' <remarks></remarks>
        Public Shared Sub Logoff(ByVal Pagina As Page)

            ' limpa authenticações
            FormsAuthentication.SignOut()

            ' registra um usuário vazio
            Dim TckNome As String = ""
            Dim TckAuth As New FormsAuthenticationTicket(1, TckNome, Now, Now.AddHours(2), False, TckNome)
            Dim TckEncr As String = FormsAuthentication.Encrypt(TckAuth)
            Dim HttpC As HttpCookie = New HttpCookie(FormsAuthentication.FormsCookieName, TckEncr)
            Pagina.Response.Cookies.Add(HttpC)
            Pagina.Session.Timeout = 120

            ' apaga registro de usuário
            Dim NaSessao As LogonSession = Logon(Pagina)
            If Not IsNothing(NaSessao) Then
                Logon(Pagina) = Nothing
            End If

            ' volta para tela de login
            FormsAuthentication.RedirectToLoginPage()
        End Sub

        Public Shared Function MasterAcessoOK(ByVal Page As Page, ByVal StrConn As String, ByVal Esquema As String) As Boolean
            Dim Ret As Boolean = False
            Dim Logins As Object = Form.BuscaTipo(Page.Master, "ASP.uc_icftlogin_icftlogin_ascx")
            Dim Login As Object = Nothing
            For Each L As Object In Logins
                If L.tipo.ToString = "PopupLogin" Then
                    Login = L
                    Exit For
                End If
            Next

            If Not IsNothing(Login) Then
                Ret = AcessoOK(Page, StrConn, Esquema, "", "")
                If Not Ret Then
                    Login.Mostra()
                End If
            Else
                Ret = AcessoOK(Page, StrConn, Esquema, "", "")
            End If
            Return Ret
        End Function

        Public Shared Function AcessoOK(ByVal Page As Page, ByVal StrConn As String, ByVal Esquema As String, ByVal UrlLogin As String, ByVal UrlRedirOK As String, ByVal ParamArray Params() As Object) As Boolean
            Try

                ' pega usuário e senha que foram passados (caso tenha sido passados)
                Dim Acesso As Boolean = False
                Dim Usu As String = ""
                Dim Snh As String = ""
                Try
                    Usu = StrStr(MacroSubstSQLText("[:usuario]", Params), 1, -1)
                    Snh = StrStr(MacroSubstSQLText("[:senha]", Params), 1, -1)
                Catch
                End Try

                ' se não forem informados, irá considerar usuário e senha já logados, caso seja possível
                Try
                    If NZ(Usu, "") = "" Then
                        Usu = Icraft.IcftBase.Logon(Page).Usuario
                        Snh = Icraft.IcftBase.DecrypB(Icraft.IcftBase.Logon(Page).Senha)
                    End If
                Catch
                End Try

                ' monta conexão com usuário e senha obtido e testa na base
                Try
                    Dim Conn As System.Configuration.ConnectionStringSettings = StrConnObj(StrConn, ";user:" & Usu & ";password:" & Snh)
                    If Conn.ProviderName <> Icraft.IcftBase.MSAccess Then
                        ' caso seja oracle ou mysql, basta verificar acesso
                        If Not IsNothing(DSValor("count(*)", IIf(Esquema <> "", Esquema & ".", "") & "sys_config_global", StrConn, "")) Then
                            Acesso = True
                        End If
                    Else
                        ' caso seja access, pesquisa em tabela de usuário
                        Dim DS As System.Data.DataSet = DSCarrega("select usuario, senha from " & IIf(Esquema <> "", Esquema & ".", "") & "ger_usuario where ucase(usuario)=:usuario and (valido_ate is null or valido_ate>=now)", Conn, ":usuario", Usu.ToUpper)
                        If DS.Tables(0).Rows.Count > 0 Then
                            If Icraft.IcftBase.DecrypB(DS.Tables(0).Rows(0)("senha")) = Snh Then
                                Acesso = True
                            End If
                        End If
                    End If
                Catch
                End Try

                ' finalmente, registra acesso caso este seja permitido e não tenha sido registrado ainda
                If Acesso Then
                    If Icraft.IcftBase.Logon(Page).Usuario <> Usu Or Icraft.IcftBase.Logon(Page).Senha <> Snh Then
                        Icraft.IcftBase.LoginCertifica(Page, Usu, Icraft.IcftBase.EncrypB(Snh), False)
                        Dim z As Integer = 0
                        Do While z < Params.Count
                            If Not Icraft.IcftBase.Compare(Params(z), ":usuario", True) AndAlso Not Icraft.IcftBase.Compare(Params(z), ":senha", True) Then
                                Logon(Page).ExtendedProps(Params(z)) = Params(z + 1)
                            End If
                            z += 2
                        Loop
                    End If
                    If NZ(UrlRedirOK, "") <> "" AndAlso Page.Request.Path <> UrlRedirOK Then
                        Page.Response.Redirect(UrlRedirOK)
                    End If
                    Return True
                End If
            Catch
            End Try

            ' caso sem permissão, direciona para rotina de acesso...
            If NZ(UrlLogin, "") <> "" AndAlso Page.Request.Path <> UrlLogin Then
                Page.Response.Redirect(UrlLogin & "?returnurl=" & UrlRedirOK)
            End If
            Return False
        End Function


        Public Shared Property LinhaQuantidade(ByVal MyBaseInst As Control) As Integer
            Get
                Return PropE(MyBaseInst, "Linhas_Quantidade")
            End Get
            Set(ByVal value As Integer)
                PropE(MyBaseInst, "Linhas_Quantidade") = value
            End Set
        End Property
        Public Shared Sub LinhaInvalidaTodas(ByVal MyBaseInst As Control)
            For z As Integer = 1 To LinhaQuantidade(MyBaseInst)
                LinhaInvalida(MyBaseInst, z) = True
            Next
        End Sub
        Public Shared Sub LinhaSalvaRegNoBuffer(ByVal MyBaseInst As Control, ByVal Container As Object, ByVal Prefixo As String, Optional ByVal Tipo As String = "", Optional ByVal NumLinha As Integer = -1)
            Dim log As String = ""
            ' -1 para salvar todas as linhas
            Dim Ini As Integer = NumLinha, Fim As Integer = NumLinha
            If NumLinha = -1 Then
                Ini = 1
                Fim = LinhaQuantidade(MyBaseInst)
            End If
            For z As Integer = Ini To Fim
                Dim DS As DataSet = New DataSet
                For Each ctl As Object In Form.Controles(Container, Prefixo)
                    Dim NomeCtl As String = Mid(ctl.ID, 4)
                    If DS.Tables(0).Rows.Count = 0 Then
                        DS.Tables(0).Rows.Add()
                    End If
                    If Not DS.Tables(0).Columns.Contains(NomeCtl) Then
                        DS.Tables(0).Columns.Add(NomeCtl)
                    End If
                    DS.Tables(0).Rows(0).Item(NomeCtl) = Controle.ValorAtual(ctl)
                Next
                LinhaDS(MyBaseInst, z, Tipo) = DS
            Next
        End Sub
        Public Shared Property LinhaRegistroNovo(ByVal MyBaseInst As Control, ByVal NumLinha As Integer) As Boolean
            Get
                Return Mid(PropE(MyBaseInst, "Linha_Nova"), NumLinha, 1) = "X"
            End Get
            Set(ByVal value As Boolean)
                Mid(PropE(MyBaseInst, "Linha_Nova"), NumLinha, 1) = IIf(value, "X", " ")
            End Set
        End Property
        Public Shared Property LinhaInvalida(ByVal MyBaseInst As Control, ByVal NumLinha As Integer) As Boolean
            Get
                Dim P As String = PropE(MyBaseInst, "Linha_Inv")
                If NumLinha > Len(P) Then
                    Return True
                End If
                Return (Mid(P, NumLinha, 1) = "X")
            End Get
            Set(ByVal value As Boolean)
                Dim P As String = PropE(MyBaseInst, "Linha_Inv")
                If Len(P) < NumLinha Then
                    P &= New String("X", NumLinha - Len(P))
                End If
                Mid(P, NumLinha, 1) = IIf(value, "X", " ")
                PropE(MyBaseInst, "Linha_Inv") = P
            End Set
        End Property
        Public Shared Property LinhaDS(ByVal MyBaseInst As Control, ByVal NumLinha As Integer, Optional ByVal Tipo As String = "") As DataSet
            Get
                Return PropE(MyBaseInst, "Linha_" & NumLinha & "_DS" & Tipo)
            End Get
            Set(ByVal value As DataSet)
                PropE(MyBaseInst, "Linha_" & NumLinha & "_DS" & Tipo) = value
            End Set
        End Property
        Public Shared Property LinhaSelecionada(ByVal MyBaseInst As Control, ByVal NumLinha As Integer) As Boolean
            Get
                Return Mid(PropE(MyBaseInst, "Linha_Sel"), NumLinha, 1) = "X"
            End Get
            Set(ByVal value As Boolean)
                Mid(PropE(MyBaseInst, "Linha_Sel"), NumLinha, 1) = IIf(value, "X", " ")
            End Set
        End Property
        Public Shared Function LinhaAlterada(ByVal MyBaseInst As Control, ByVal NumLinha As Integer, ByVal Tipo As String, ByVal TipoCompara As String, Optional ByVal NumLinhaCompara As Integer = -1) As Boolean
            If NumLinhaCompara = -1 Then
                NumLinhaCompara = NumLinha
            End If
            Dim DS As DataSet = LinhaDS(MyBaseInst, NumLinha, Tipo)
            Dim DSCompara As DataSet = LinhaDS(MyBaseInst, NumLinhaCompara, TipoCompara)
            For Each CampoCtl As DataColumn In DS.Tables(0).Columns
                If DS.Tables(0).Rows(0).Item(CampoCtl) <> DSCompara.Tables(0).Rows(0).Item(CampoCtl) Then
                    Return True
                End If
            Next
            Return False
        End Function
        Public Shared Function LinhaAlterada(ByVal MyBaseInst As Control, ByVal Container As Object, ByVal Prefixo As String, ByVal NumLinhaCompara As Integer, Optional ByVal TipoCompara As String = "") As Boolean
            Dim DSCompara As DataSet = LinhaDS(MyBaseInst, NumLinhaCompara, TipoCompara)
            For Each CampoCtl As Control In Form.Controles(Container, Prefixo)
                Dim NomeCtl As String = Mid(CampoCtl.ID, 4)
                If Prop(CampoCtl) <> DSCompara.Tables(0).Rows(0).Item(NomeCtl) Then
                    Return True
                End If
            Next
            Return False
        End Function



        Shared Function TextoLogEx(ByVal ObjErr As Exception, Optional ByVal TextoSimplif As String = "") As String
#If _MYTYPE = "WindowsForms" Then

            Dim Err As String = "----------------------------------------------------------------------------------------" & vbCrLf & "Momento: " & Format(Now, "yyyy-MM-dd ddd HH:mm:ss") & vbCrLf
            Err &= "Produto: " & Application.ProductName & vbCrLf
            Err &= "Aplicativo: " & Application.ExecutablePath & vbCrLf & vbCrLf
            If TextoSimplif <> "" Then
                Err &= "Mensagem apresentada: " & TextoSimplif & vbCrLf & vbCrLf
            End If

            If Not IsNothing(ObjErr) Then


                Err &= "Mensagem de erro:" & vbCrLf
                Err &= ObjErr.Message & vbCrLf & vbCrLf
                If Not IsNothing(ObjErr.StackTrace) Then
                    Err &= "Erro:" & vbCrLf
                    Err &= ObjErr.StackTrace.ToString().Replace(Chr(9), "    ")
                    Err &= vbCrLf & vbCrLf
                End If

            End If

#Else
            Dim ctx As HttpContext = HttpContext.Current
            Dim err As String = "<span style='font-family:Verdana;font-weight:normal;font-size:.7em;color:black'>"
            err &= "<span style='font-size:18pt;color:red'>Erro no site " & ctx.Server.HtmlEncode(WebConf("site_nome")) & "</span><br />"
            err &= "<b>Momento:</b> " & Format(Now, "yyyy-MM-dd ddd HH:mm:ss") & "<br /><br />"
            err &= "<b>P&aacute;gina de erro:</b> " & ctx.Request.Url.ToString() & "<br /><br />"
            err &= "<b>IP do host:</b> " & ctx.Request.UserHostAddress.ToString & "<br /><br />"

            Try
                If ctx.Session.Keys.Count > 0 Then
                    err &= "<b>Vari&aacute;veis de sess&atilde;o:</b><ul>"
                    For Each var As String In ctx.Session.Keys
                        err &= "<li><b>" & ctx.Server.HtmlEncode(var) & "</b> = "
                        Try
                            err &= NZ(ctx.Session(var), "")
                        Catch
                            err &= ctx.Session(var).GetType.ToString()
                        End Try
                        err &= "</li>"
                    Next
                    err &= "</ul>"
                End If
                err &= "<br />"
            Catch
            End Try

            err &= "<b>Servidor:</b> " & ctx.Server.MachineName & "<br /><br />"
            If TextoSimplif <> "" Then
                err &= "<b>Mensagem apresentada: " & ctx.Server.HtmlEncode(TextoSimplif) & "</b><br /><br />"
            End If
            If Not IsNothing(ObjErr) Then
                err &= "<b>Mensagem de erro:</b> "
                If TypeOf (ObjErr) Is HttpException Then
                    With CType(ObjErr, HttpException)
                        err &= .GetHttpCode & " - " & ctx.Server.HtmlEncode(ObjErr.Message)
                    End With
                Else
                    err &= ctx.Server.HtmlEncode(ObjErr.Message)
                End If
                err &= "<br /><br />"
                If Not IsNothing(ObjErr.StackTrace) Then
                    err &= "<b>Erro:</b><br /><span style='font-size:12px'>"
                    err &= ObjErr.StackTrace.ToString().Replace(Chr(13) & Chr(10), "<br />").Replace(Chr(9), "    ")
                    err &= "</span><br /><br />"
                End If
            End If
            err &= "<hr />"
            err &= "<br /></span>"
#End If
            Return err
        End Function
        Shared Sub ErroLogReg(ByVal Ex As Exception, Optional ByVal TextoSimplif As String = "")
            Dim Texto As String = TextoLogEx(Ex, TextoSimplif)
            If NZ(WebConf("erro_remetente"), "") <> "" AndAlso WebConf("erro_notifica") <> "" Then
                EnviaEmail(WebConf("erro_remetente"), WebConf("erro_notifica"), "Erro no site " & WebConf("site_nome"), Texto, Net.Mail.MailPriority.High)
            End If
            If NZ(WebConf("erro_arqlog"), "") <> "" Then
                GravaLog(WebConf("erro_arqlog"), Texto)
            End If
        End Sub

        ''' <summary>
        ''' Retorna mensagem tratada evitando dados técnicos em apresentação para usuário.
        ''' </summary>
        ''' <param name="Ex">Exceção a ser tratada.</param>
        ''' <param name="MensagemCompl">Texto de introdução. Será incluído no início da mensagem.</param>
        ''' <returns>Retorna texto a ser apresentado para o usuário, considerando o tratamento de erros previstos.</returns>
        ''' <remarks>Veja o padrão. A substituição ocorre sem ponto no final, por favor.</remarks>
        Public Shared Function MessageEx(ByVal Ex As Exception, Optional ByVal MensagemCompl As String = "") As String

            ' mensagem padrão
            Dim Mensagem As String = Ex.Message

            If Not IsNothing(Ex.InnerException) AndAlso NZ(Ex.InnerException.Message, "") <> "" Then
                Mensagem &= ". " & Ex.InnerException.Message
            End If
            Dim Param As String

            ' mensagens específicas
            Param = RegexGroup(Mensagem, "Cannot update (.*); field not updateable", 1).Value
            If Param <> "" Then
                Mensagem = "Por restrições da base de dados, campo " & Param & " não pode ser atualizado"
            End If

            Param = RegexGroup(Mensagem, "create duplicate values in the").Value
            If Param <> "" Then
                Mensagem = "Tentativa de registro de chave duplicada"
            End If

            Param = RegexGroup(Mensagem, "Cannot set column (.*). The value violates the MaxLength.*", 1).Value
            If Param <> "" Then
                Mensagem = "Tamanho do campo " & Param & " excede o limite"
            End If

            Param = RegexGroup(Mensagem, "The path is not of a legal").Value
            If Param <> "" Then
                Mensagem = "Caminho de arquivo inexistente ou ilegal"
            End If

            Param = RegexGroup(Mensagem, "Duplicate entry (.*) for key .*", 1).Value
            If Param <> "" Then
                Mensagem = "Tentativa de gravação de registro duplicado - " & Param
            End If

            Param = RegexGroup(Mensagem, "Empty path name is not legal").Value
            If Param <> "" Then
                Mensagem = "Nome de arquivo incorreto"
            End If

            Param = RegexGroup(Mensagem, "Could not find file '(.*?)'", 1).Value
            If Param <> "" Then
                Mensagem = "Arquivo não encontrado: " & Param
            End If


            ' ------------------------------------------------
            ' TRATAMENTO DE ERROS DO ORACLE

            Param = RegexGroup(Mensagem, "ORA-02291: .*\((.*)\)", 1).Value
            If Param <> "" Then
                Mensagem = "Falta de registro relacionado em " & Param
            End If

            Param = RegexGroup(Mensagem, "ORA-00001: .*\((.*)\)", 1).Value
            If Param <> "" Then
                Mensagem = "Tentativa de registro de chave duplicada em " & Param
            End If

            Param = RegexGroup(Mensagem, "ORA-01017:").Value
            If Param <> "" Then
                Mensagem = "Logon incorreto. Usuário ou senha inválidos ou sessão expirada"
            End If

            Param = RegexGroup(Mensagem, "ORA-00942:").Value
            If Param <> "" Then
                Mensagem = "Tabela ou visão inexistente"
            End If

            Param = RegexGroup(Mensagem, "ORA-12541:|ORA-12170:").Value
            If Param <> "" Then
                Mensagem = "Banco de dados indisponível no momento. Suporte já foi contactado"
            End If

            ' ------------------------------------------------
            ' TRATAMENTO DE ERROS DO MYSQL

            Param = RegexGroup(Mensagem, "Access denied for user (.*)", 1).Value
            If Param <> "" Then
                Mensagem = "Acesso não autorizado para " & Param & ". Verifique usuário e senha e tente novamente"
            End If

            Mensagem = IIf(MensagemCompl <> "", MensagemCompl & ". ", "") & Mensagem & "."
            Return Mensagem
        End Function


        ''' <summary>
        ''' Retorna a quantidade de estruturas regex encontradas na procura do select.
        ''' </summary>
        ''' <param name="Campos">Lista de campos onde os critérios serão procurados.</param>
        ''' <param name="Criterios">Critérios regex separados por ponto e vírgula.</param>
        ''' <returns>Retorna uma expressão capaz de mencionar quantidade de ocorrências do critério no campo.</returns>
        ''' <remarks></remarks>
        Shared Function RegExpCount(ByVal Campos As String, ByVal Criterios As String) As String
            Dim Campo As String = Split(Campos, ";")(0)
            Dim Expr As String = "(" & Criterios.Replace(" ", "|") & ")"
            Return "(SELECT COUNT(LEVEL) AS QTD FROM DUAL CONNECT BY REGEXP_INSTR(" & Campo & ", " & SqlExpr(Expr) & ",1,LEVEL)>0)"
        End Function

        ''' <summary>
        ''' Retorna posição de texto em campo com base em cláusula select preparado para cláusula select.
        ''' </summary>
        ''' <param name="Campos">Nome do campo de pesquisa.</param>
        ''' <param name="Criterios">Critério regex a ser aplicado.</param>
        ''' <returns>Retorna texto pronto para ser colocado como campo do select.</returns>
        ''' <remarks></remarks>
        Shared Function RegExpInstr(ByVal Campos As String, ByVal Criterios As String) As String
            Dim Campo As String = Split(Campos, ";")(0)
            Dim Criterio As String = Criterios.Replace(" ", "|")
            Return "REGEXP_INSTR(" & Campo & ",'" & Criterio & "')"
        End Function

        ''' <summary>
        ''' Retorna campo pronto para cláusula select com amostra do registro encontrado baseada em código regex.
        ''' </summary>
        ''' <param name="Campos">Nome do campo de pesquisa.</param>
        ''' <param name="Criterios">Critério regex a ser aplicado.</param>
        ''' <param name="Destaque">Formato da string de retorno considerando "\1" para grupo 1 e assim por diante.</param>
        ''' <returns>Retorna texto pronto para ser colocado como campo do select.</returns>
        ''' <remarks></remarks>
        Shared Function RegExpAmostra(ByVal Campos As String, ByVal Criterios As String, Optional ByVal Destaque As String = "<span style='background-color:yellow'>\1</span>\2", Optional ByVal QtdAntesouDepois As Integer = 15) As String
            Dim Campo As String = Split(Campos, ";")(0)
            Dim Criterio As String = Criterios.Replace(" ", "|")
            Return "REGEXP_REPLACE(TRANSLATE(SUBSTR(" & Campo & ", GREATEST(" & RegExpInstr(Campos, Criterios) & "-" & QtdAntesouDepois & ",1) ),'<>','  '),'(" & Criterio & ")(.{0," & QtdAntesouDepois & "}).*', '" & Destaque.Replace("'", "''") & "')"
        End Function

        ''' <summary>
        ''' Cria expressão para procura por ocorrência de expressão regex entre os campos informados. 
        ''' </summary>
        ''' <param name="Campos">Campos a serem tratados.</param>
        ''' <param name="Criterios">Critérios regex a serem aplicados ao mesmo tempo. Ex: "[cC][aA][sS][aA] [vV][eE][lL][hH][aA]" (texto com os dois ao mesmo tempo).</param>
        ''' <returns>Retorna sequência de select para pesquisa considerando condições regex.</returns>
        ''' <remarks></remarks>
        Shared Function RegExpLike(ByVal Campos As String, ByVal Criterios As String) As String
            Dim CondOr As String = ""
            For Each Campo As String In Split(Campos, ";")
                Dim CondAnd As String = ""
                For Each Pedaco As String In Split(Criterios, " ")
                    CondAnd &= IIf(CondAnd <> "", " AND ", "") & "REGEXP_LIKE(" & Campos & ", " & SqlExpr(Pedaco) & ")"
                Next
                CondOr &= IIf(CondOr <> "", " OR ", "") & "(" & CondAnd & ")"
            Next
            Return CondOr
        End Function

        ''' <summary>
        ''' Retorna expressão de pesquisa REGEX baseando-se na string passada.
        ''' </summary>
        ''' <param name="TEXTO">Texto para origem da expressão.</param>
        ''' <returns>Código de pesquisa desconsiderando os acentos.</returns>
        ''' <remarks></remarks>
        Shared Function RegExpSemAcento(ByVal TEXTO As String) As String
            Dim Ret As String = ""
            For i As Integer = 1 To Len(TEXTO)
                Dim Letra As String = Mid(TEXTO, i, 1)
                Select Case Letra
                    Case "a", "A", "á", "Á", "à", "À", "ã", "Ã", "â", "Â", "â", "ä", "Ä"
                        Letra = "[áÁàÀãÃâÂâäÄaA]"
                    Case "e", "E", "é", "É", "ê", "Ê", "Ë", "ë", "È", "è"
                        Letra = "[éÉêÊËëÈèeE]"
                    Case "i", "I", "í", "Í", "ï", "Ï", "Ì", "ì"
                        Letra = "[íÍïÏÌìiI]"
                    Case "o", "O", "ó", "Ó", "ô", "Ô", "õ", "Õ", "ö", "Ö", "ò", "Ò"
                        Letra = "[óÓôÔõÕöÖòÒoO]"
                    Case "u", "U", "ú", "Ú", "Ù", "ù", "ú", "û", "ü", "Ü", "Û"
                        Letra = "[úÚÙùúûüÜÛuU]"
                    Case "c", "C", "ç", "Ç"
                        Letra = "[çÇcC]"
                    Case "n", "N", "ñ", "Ñ"
                        Letra = "[nNñÑ]"
                    Case Else
                        If LCase(Letra) <> UCase(Letra) Then
                            Letra = "[" & LCase(Letra) & UCase(Letra) & "]"
                        End If
                End Select
                Ret &= Letra
            Next
            Return Ret
        End Function



        ' -----------------------------------------------------------------
        ' obj obsoleto, utilizado para compatibilidade de versão


        Shared Property Def(ByVal Controles As Control, ByVal NomeControle As String, Optional ByVal Propriedade As String = "") As Object
            Get
                Dim ctl As Control = Controles.FindControl(NomeControle)
                Return Prop(ctl, Propriedade)
            End Get
            Set(ByVal value As Object)
                Dim ctl As Control = Controles.FindControl(NomeControle)
                Prop(ctl, Propriedade) = value
            End Set
        End Property



        Shared Function IncluiCampo(ByVal Objeto As Object, ByVal Container As Table, ByVal Prefixo As String, ByVal Nome As String, ByVal LarguraCampo As String, ByVal Tipo As String, ByVal Etiqueta As String, ByVal LarguraEtiq As String, ByVal ToolTip As String, ByVal ExtendedProps As String, ByVal Formato As String, ByVal Tamanho As String, ByVal Auto As String, ByVal ValorPadrao As String, ByVal Sistema As String, ByVal Tabela As String, ByVal StrGerador As String, ByVal StrConn As String, ByVal Bloqueado As Boolean, Optional ByVal Estrut As Object = Nothing) As Control
            Dim Page As Page
            If TypeOf Objeto Is Page Then
                Page = Objeto
                Objeto = Nothing
            Else
                Page = Objeto.Page
            End If
            Dim etiq As Panel = Nothing
            Dim lbl As Label

            If Not Nome Like "SYS_*" Then



                ' disponibiliza propriedades do gerador...
                Dim Props As ElementosStr = New ElementosStr(ExtendedProps)

                ' verifica formato
                If Formato = "" Then
                    If Compare(Tipo, "System.Boolean") Then
                        Formato = "BOOL"
                    ElseIf Compare(Tipo, "System.Decimal") Then
                        Formato = "CURRENCY"
                    ElseIf Compare(Tipo, "System.Byte") Or Compare(Tipo, "System.Int32") Then
                        Formato = "INTEGER"
                    ElseIf Compare(Tipo, "System.Double") Or Compare(Tipo, "System.Single") Then
                        Formato = "REAL"
                    ElseIf Compare(Tipo, "System.DateTime") Then
                        Formato = "dd\/MM\/yyyy"
                    End If
                End If

                Dim trow As New TableRow



                Dim tcel As New TableCell
                tcel.VerticalAlign = VerticalAlign.Top

                ' inclui etiqueta
                If Etiqueta <> "" Then
                    etiq = New Panel
                    etiq.Style("text-align") = "left"
                    etiq.Style("float") = "left"
                    etiq.Style("margin-left") = "20px"
                    etiq.ID = "divlbl" & Nome
                    etiq.CssClass = "icftform_etiq"
                    etiq.Style.Add("width", LarguraEtiq & "px")
                    etiq.ToolTip = ToolTip

                    Dim etiqlbl As New Label
                    etiqlbl.ID = "lbl" & Nome
                    etiqlbl.Text = Etiqueta

                    etiq.Controls.Add(etiqlbl)
                    tcel.Controls.Add(etiq)
                End If

                trow.Cells.Add(tcel)

                tcel = New TableCell

                ' campo, caso seja booleano <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                If Compare(Formato, "bool") Then
                    Dim bool As CheckBox

                    bool = New CheckBox
                    bool.ID = Prefixo & Nome
                    bool.CssClass = "icftform_checkbox"
                    bool.ToolTip = ToolTip
                    bool.Attributes("ValorPadrao") = ValorPadrao
                    bool.Style.Add("position", "relative")
                    bool.Style.Add("left", -(LarguraCampo / 2) - 5 & "px")

                    ' define outras propriedades
                    For Each prop As ElementoStr In Props.Elementos
                        bool.Style.Add(prop.Nome, prop.Conteudo)
                    Next

                    tcel.Controls.Add(bool)

                    lbl = New Label
                    lbl.ID = "lblbr" & Nome
                    lbl.Text = "<br clear='all'/>"
                    tcel.Controls.Add(lbl)


                    trow.Cells.Add(tcel)
                    trow.Cells(0).Style("border-bottom") = "1px solid #f0f0f0"
                    Container.Rows.Add(trow)
                    If Not IsNothing(etiq) Then
                        bool.Attributes("Etiq") = etiq.ID
                    End If


                    Return bool
                End If

                ' campo, caso seja html <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                If Compare(Formato, "html") Then
                    Dim div As New Panel
                    div.Style("padding-bottom") = "60px"
                    div.Style("display") = "inline-block"

                    Dim ctl As Object = Page.LoadControl("~\uc\icfttextarea\icfttextarea.ascx")
                    ctl.ID = Prefixo & Nome
                    ctl.Panel.CssClass = "icftform_html"
                    ctl.Largura = LarguraCampo & "PX"
                    ctl.Attributes("ValorPadrao") = ValorPadrao

                    ' define outras propriedades
                    For Each prop As ElementoStr In Props.Elementos
                        ctl.estilo(prop.Nome) = prop.Conteudo
                    Next

                    ctl.DATABIND()
                    div.Controls.Add(ctl)

                    trow.Cells.Add(New TableCell)
                    Container.Rows.Add(trow)

                    trow = New TableRow
                    tcel = New TableCell
                    tcel.Controls.Add(div)
                    tcel.ColumnSpan = 2
                    tcel.Style("text-align") = "right"

                    trow.Cells.Add(tcel)
                    trow.Cells(0).Style("border-bottom") = "1px solid #f0f0f0"
                    Container.Rows.Add(trow)
                    If Not IsNothing(etiq) Then
                        ctl.Attributes("Etiq") = etiq.ID
                    End If


                    Return ctl
                End If

                ' combo, se com relacionamento ou opções no ToolTip <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<,
                Dim toolTipMatches As MatchCollection = RegularExpressions.Regex.Matches(NZ(ToolTip, ""), "\b(\w+)=(\w+)\b")
                If (Sistema <> "" And Tabela <> "") OrElse toolTipMatches.Count > 0 Then
                    ' busca relacionamento n1

                    Dim rels As New RelsN1(Sistema, Tabela, Nome, StrGerador)

                    If rels.Count <> 0 OrElse toolTipMatches.Count > 0 Then
                        Dim list As DropDownList

                        list = New DropDownList
                        list.ID = Prefixo & Nome
                        list.ToolTip = ToolTip
                        list.CssClass = "icftform_lista"
                        list.Attributes("ValorPadrao") = ValorPadrao
                        list.Style("width") = LarguraCampo & "px"

                        Dim Atualiza As ArrayList = New ArrayList

                        If rels.Count <> 0 Then
                            For z As Integer = 0 To rels.Count - 1
                                If rels(z)("_CAMPOITEM") < (rels(z)("_CAMPOSQTD") - 1) Then
                                    Dim Campos As Array = Split(rels(z)("CAMPO_N"), ";")
                                    For zz As Integer = rels(z)("_CAMPOITEM") + 1 To rels(z)("_CAMPOITEM") + 1
                                        If Not Atualiza.Contains(Campos(zz)) Then
                                            Atualiza.Add(Prefixo & Campos(zz))
                                        End If
                                    Next
                                    Exit For
                                End If
                            Next
                            If Atualiza.Count > 0 Then
                                list.AutoPostBack = True
                                PropE(list, "Atualizar") = Join(Atualiza.ToArray, ";")
                                AddHandler list.TextChanged, AddressOf AtualizouControle

                                If Not IsNothing(Estrut) Then
                                    Estrut.CodInit.AppendLine("AddHandler list.TextChanged, AddressOf AtualizouControle")
                                End If

                            End If

                            Dim sql As String = ""
                            If rels(0)("_CAMPOITEM") > 0 Then
                                Dim REL1 As Array = Split(rels(0)("CAMPO_1"), ";")
                                Dim RELN As Array = Split(rels(0)("CAMPO_N"), ";")
                                For Z As Integer = 0 To rels(0)("_CAMPOITEM") - 1
                                    sql = sql & IIf(sql <> "", " AND ", "") & REL1(Z) & " = [:" & Prefixo & RELN(Z) & "]"
                                Next
                            End If

                            Dim chave_apres As String = NZ(rels(0)("chave_apres_1"), "")
                            If chave_apres = "" Then
                                chave_apres = NZ(rels(0)("chave_apres"), "")
                            End If
                            Dim chave_apres_Array() As String = {}
                            If chave_apres <> "" Then
                                chave_apres_Array = Split(chave_apres, ";")
                                chave_apres = ", " & Join(chave_apres_Array, ", ")
                            End If

                            sql = "select " & rels(0)("_CAMPOREL") & chave_apres & " from " & rels(0)("TABELA_1").ToString.ToLower & IIf(sql <> "", " WHERE ", "") & sql
                            PropE(list, "SQL") = sql
                            PropE(list, "StrConn") = StrConn
                            PropE(list, "QtdCols") = 1 + chave_apres_Array.Length
                        Else
                            For Each mt As Match In toolTipMatches
                                If Not String.IsNullOrEmpty(mt.Groups(0).Value) Then
                                    Atualiza.Add(mt.Groups(1).Value)
                                    Atualiza.Add(mt.Groups(2).Value)
                                End If
                            Next

                            If Atualiza.Count > 0 Then
                                list.AutoPostBack = True
                                PropE(list, "Atualizar") = Join(Atualiza.ToArray, ";")
                                AddHandler list.TextChanged, AddressOf AtualizouControle
                            End If

                            CarregaCombo(list, 2, False, " | ", True, Atualiza)
                        End If


                        ' define outras propriedades
                        For Each prop As ElementoStr In Props.Elementos
                            list.Style.Add(prop.Nome, prop.Conteudo)
                        Next


                        tcel.Controls.Add(list)

                        ' inclui quebra de linha
                        lbl = New Label
                        lbl.ID = "lblbr" & Nome
                        lbl.Text = "<br clear='all'/>"
                        tcel.Controls.Add(lbl)


                        trow.Cells.Add(tcel)
                        trow.Cells(0).Style("border-bottom") = "1px solid #f0f0f0"
                        Container.Rows.Add(trow)
                        If Not IsNothing(etiq) Then
                            list.Attributes("Etiq") = etiq.ID
                        End If

                        Return list
                    End If
                End If

                ' outros campos <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                Dim txt As TextBox
                txt = New TextBox
                txt.ID = Prefixo & Nome

                If Compare(Tipo, "System.DateTime") Then
                    txt.Style.Add("width", LarguraCampo - 10 & "px")
                Else
                    txt.Style.Add("width", LarguraCampo & "px")
                End If

                If Compare(Formato, "memo") Then
                    txt.CssClass = "icftform_memo"
                    txt.TextMode = TextBoxMode.MultiLine
                Else
                    txt.CssClass = "icftform_txt"
                End If

                txt.ToolTip = ToolTip
                txt.Attributes("Formato") = Formato
                txt.Attributes("Auto") = Auto
                txt.Attributes("ValorPadrao") = ValorPadrao
                txt.Attributes("Tipo") = Tipo

                If Auto <> "" Then
                    txt.Attributes("Tabela") = Tabela
                    txt.Attributes("Campo") = Nome
                    txt.Attributes("StrConn") = StrConn
                End If

                Dim Tam As Integer = NZV(Tamanho, "0")
                If Tam <> 0 Then
                    txt.MaxLength = Tamanho
                    Prop(txt, "ToolTip") &= " [Máximo de " & txt.MaxLength & " caracter" & IIf(txt.MaxLength > 1, "es", "") & "]"
                End If

                If Compare(Formato, "senha") Then
                    txt.TextMode = TextBoxMode.Password
                End If


                ' inclui o campo
                txt.Enabled = Not Bloqueado

                ' define outras propriedades
                For Each prop As ElementoStr In Props.Elementos
                    txt.Style.Add(prop.Nome, prop.Conteudo)
                Next
                tcel.Controls.Add(txt)

                ' inclui botão do calendário
                If (Compare(Formato, "dd\/MM\/yyyy")) Then
                    Dim btnCalend As New HyperLink
                    btnCalend.ID = "btnCalenda" & Nome
                    btnCalend.Text = "+"
                    btnCalend.NavigateUrl = "javascript:void(0);"
                    btnCalend.Attributes.Add("onclick", "displayCalendar($_('" & txt.ID & "',this),'dd/mm/yyyy',this)")
                    tcel.Controls.Add(btnCalend)

                    IncluiScript(Page, "js1_calendario", "<script>var host = '" & Page.ResolveUrl(NZV(WebConf("url_site"), "~/")) & "';var pathToImages = host + 'uc/IcftCalendario/Theme/" & NZV(WebConf("Theme"), "Default") & "/images/';</script>")
                    IncluiScript(Page, "js2_calendario", Page.ResolveUrl("~/uc/IcftCalendario/Theme/" & NZV(WebConf("Theme"), "Default") & "/dhtmlgoodies_calendar.js"))
                    IncluiScript(Page, "js3_calendario", Page.ResolveUrl("~/uc/IcftCalendario/dhtmlgoodies_calendar.js"))
                    IncluiStyleSheet(Page, btnCalend.UniqueID, NZV(WebConf("calendario_css"), "~/uc/IcftCalendario/Theme/" & NZV(WebConf("Theme"), "Default") & "/dhtmlgoodies_calendar.css"))

                    If Not IsNothing(Estrut) Then
                        Estrut.CodInit.AppendLine("IncluiScript(Page, ""js1_calendario"", ""<script>var host = '"" & Page.ResolveUrl(NZV(WebConf(""url_site""), ""~/"")) & ""';var pathToImages = host + 'uc/IcftCalendario/Theme/"" & NZV(WebConf(""Theme""), ""Default"") & ""/images/';</script>"")")
                        Estrut.CodInit.AppendLine("IncluiScript(Page, ""js2_calendario"", Page.ResolveUrl(""~/uc/IcftCalendario/Theme/"" & NZV(WebConf(""Theme""), ""Default"") & ""/dhtmlgoodies_calendar.js""))")
                        Estrut.CodInit.AppendLine("IncluiScript(Page, ""js3_calendario"", Page.ResolveUrl(""~/uc/IcftCalendario/dhtmlgoodies_calendar.js""))")
                        Estrut.CodInit.AppendLine("IncluiStyleSheet(Page, """ & btnCalend.UniqueID & """, NZV(WebConf(""calendario_css""), ""~/uc/IcftCalendario/Theme/"" & NZV(WebConf(""Theme""), ""Default"") & ""/dhtmlgoodies_calendar.css""))")
                    End If

                End If


                ' imagem
                Dim Reg As Match = Regex.Match(Formato, "(?is)ARQUIVO($|;)(.*)")
                If Reg.Groups(0).Value <> "" Then
                    Dim Elem As New ElementosStr(Reg.Groups(2).Value)
                    Dim Div As New Panel
                    Div.ID = "pnlDialog" & Nome
                    Div.CssClass = "icftform_comandos"
                    Dim ctl As Object = Page.LoadControl("~\uc\icftdialogo\icftdialogo.ascx")
                    ctl.id = "dlgNovo" & Nome
                    ctl.tipo = 1
                    ctl.controlevinc = txt.ID
                    ctl.mascara = NZV(Elem.Items("mascara").Conteudo, "*.*")
                    ctl.caminho = NZV(Elem.Items("caminho").Conteudo, "~/img")
                    If NZV(Elem.Items("salvasemcaminho").Conteudo, True) Then
                        ctl.OBTERTEXTOSEMCAMINHO = True
                    End If
                    ctl.EscondeTexto = True
                    ctl.BotaoTexto = "Novo Arquivo"
                    ctl.ToolTip = "Clique para inserir um arquivo novo."
                    ctl.titulo = "Novo Arquivo " & NZV(Elem.Items("mascara").Conteudo, "*.jpg")
                    ctl.PermitirAlterarCaminho = False
                    Div.Controls.Add(ctl)

                    ctl = Page.LoadControl("~\uc\icftdialogo\icftdialogo.ascx")
                    ctl.id = "dlgJaExist" & Nome
                    ctl.tipo = 2
                    ctl.controlevinc = txt.ID
                    ctl.mascara = NZV(Elem.Items("mascara").Conteudo, "*.jpg")
                    ctl.caminho = NZV(Elem.Items("caminho").Conteudo, "~/img")
                    If NZV(Elem.Items("salvasemcaminho").Conteudo, False) Then
                        ctl.OBTERTEXTOSEMCAMINHO = True
                    End If
                    ctl.EscondeTexto = True
                    ctl.BotaoTexto = "Arquivo Já existente"
                    ctl.ToolTip = "Clique para escolher um arquivo já existente no diretório."
                    ctl.titulo = "Arquivo Já Existente " & ctl.mascara
                    ctl.PermitirAlterarCaminho = False
                    Div.Controls.Add(ctl)

                    tcel.Controls.Add(Div)

                End If



                ' mascara o campo, caso seja necessário
                If Formato <> "" And Not (Compare(Formato, "HTML") Or Compare(Formato, "MEMO")) Then
                    txt.Page = Page
                    Controle.AplicaMascara(txt)
                End If

                ' inclui quebra de linha
                lbl = New Label
                lbl.ID = "lblbr" & Nome
                lbl.Text = "<br clear='all'/>"
                tcel.Controls.Add(lbl)

                trow.Cells.Add(tcel)
                trow.Cells(0).Style("border-bottom") = "1px solid #f0f0f0"
                Container.Rows.Add(trow)
                If Not IsNothing(etiq) Then
                    txt.Attributes("Etiq") = etiq.ID
                End If

                Return txt
            End If
            Return Nothing
        End Function

        Class RelsN1 ' busca relacionamentos para combobox
            Private Rels As DataView
            ReadOnly Property Count() As Integer
                Get
                    Return Rels.Table.Rows.Count
                End Get
            End Property
            Default ReadOnly Property Item(ByVal Index As Integer) As DataRowView
                Get
                    Return Rels.Item(Index)
                End Get
            End Property
            ReadOnly Property DataView() As DataView
                Get
                    Return Rels
                End Get
            End Property
            Sub New(ByVal Sistema As String, ByVal Tabela As String, ByVal Campo As String, ByVal StrGerador As String)
                Dim esquema As String = ""
                If InStr(Tabela, ".") Then
                    Dim tt As Array = Split(Tabela, ".")
                    esquema = tt(0)
                    Tabela = tt(1)
                End If
                Dim relss As DataSet = Nothing
                Try
                    'Adicionado campo chave_apres
                    'relss = DSCarrega("select r.tabela_1 as tabela_1, r.campo_1 as campo_1, r.tabela_n as tabela_n, r.campo_n as campo_n, r.obrig as obrig, r.chave_apres_1 as chave_apres_1, t1.chave_apres as chave_apres from GER_relacionamento as r, GER_tabela as t1 where r.sistema = t1.sistema and r.tabela_1 = t1.tabela and t1.sistema = :sistema and tabela_n = :tabela and (INSTR(';' & " & SWITCH() & " & ';', ';' & :campo & ';')<>0)", StrGerador, ":sistema", Sistema, ":tabela", Tabela, ":campo", UCase(Campo))

                    Select Case DSTipoBaseSQL(StrGerador)
                        Case TipoBaseSQL.MSAccess
                            relss = DSCarrega("select r.tabela_1 as tabela_1, r.campo_1 as campo_1, r.tabela_n as tabela_n, r.campo_n as campo_n, r.obrig as obrig, r.chave_apres_1 as chave_apres_1, t1.chave_apres as chave_apres from GER_relacionamento as r, GER_tabela as t1 where r.sistema = t1.sistema and r.tabela_1 = t1.tabela and t1.sistema = :sistema and tabela_n = :tabela and (INSTR(';' & UCASE(campo_n) & ';', ';' & :campo & ';')<>0)", StrGerador, ":sistema", Sistema, ":tabela", Tabela, ":campo", UCase(Campo))
                        Case Else
                            relss = DSCarrega("select r.tabela_1 tabela_1, r.campo_1 campo_1, r.tabela_n tabela_n, r.campo_n campo_n, r.obrig obrig, r.chave_apres_1 chave_apres_1, t1.chave_apres chave_apres from GER_relacionamento r, GER_tabela t1 where r.sistema = t1.sistema and r.tabela_1 = t1.tabela and t1.sistema = :sistema and tabela_n = :tabela and (INSTR(';' || UPPER(campo_n) || ';', ';' || :campo || ';')<>0)", StrGerador, ":sistema", Sistema, ":tabela", Tabela, ":campo", UCase(Campo))
                    End Select


                Catch
                End Try
                If IsNothing(relss) Then
                    relss = New DataSet
                End If
                If relss.Tables.Count = 0 Then
                    relss.Tables.Add(New DataTable)
                End If
                relss.Tables(0).Columns.Add("_CAMPOITEM", GetType(Integer))
                relss.Tables(0).Columns.Add("_CAMPOREL", GetType(String))
                relss.Tables(0).Columns.Add("_CAMPOSQTD", GetType(Integer))
                For Each row As DataRow In relss.Tables(0).Rows
                    row("_CAMPOITEM") = Array.IndexOf(Split(UCase(row("CAMPO_N")), ";"), UCase(Campo))
                    row("_CAMPOREL") = Split(UCase(row("CAMPO_1")), ";")(row("_CAMPOITEM"))
                    row("_CAMPOSQTD") = Split(row("CAMPO_N"), ";").Length
                    If esquema <> "" Then
                        row("TABELA_1") = esquema & "." & row("TABELA_1")
                        row("TABELA_N") = esquema & "." & row("TABELA_N")
                    End If
                Next
                relss.Tables(0).DefaultView.ApplyDefaultSort = True
                relss.Tables(0).DefaultView.Sort = "_CAMPOSQTD"
                Rels = relss.Tables(0).DefaultView
            End Sub
        End Class


        ''' <summary>
        ''' NotaMsg<br/>
        ''' Armazena e mostra mensagem de Nota como validador.
        ''' Para utilizar:
        ''' Try
        ''' ...
        ''' Catch ex as Exception
        ''' ...Icraft.NotaMsg.Trata(Page, Ex, IgnoreNotas, URLRedirNota)
        ''' End Try
        ''' Na página que tratará o Nota (postlocal ou redirNota), deverá colocar:
        ''' Icraft.NotaMsg.Verifica(Page) no PRE_RENDER.
        ''' </summary>
        ''' <remarks></remarks>
        ''' 
        Class NotaMsg
            <Serializable()> Class Msg
                Public Pagina As Page
                Public Controle As Control
                Public Texto As String
                Public Ex As Exception
                Public Identif As String
                Public URLRedirNota As String
                Sub New(ByVal Controle As Control, ByVal Texto As String, ByVal Nota As Exception, Optional ByVal Identif As String = "", Optional ByVal URLRedirNota As String = "")
                    Me.Pagina = Controle.Page
                    Me.Controle = Controle
                    Me.Texto = Texto
                    Me.Ex = Nota
                    Me.Identif = Identif
                    Me.URLRedirNota = URLRedirNota
                End Sub
            End Class

            ''' <summary>
            ''' Recupera ou registra dado em cache ainda para apresentação.
            ''' </summary>
            ''' <param name="Controle">Página corrente (deve ser PAGE).</param>
            ''' <value>Objeto MSG, que possui página onde msg foi gerada, exception e URL de redirecionamento.</value>
            ''' <returns>Retorna objeto MSG gerada, expception e URL de redirecionamento (página que trata do Nota).</returns>
            ''' <remarks></remarks>
            Public Shared Property Ex(ByVal Controle As Control) As Msg
                Get
                    Dim nota As Msg = Nothing
                    nota = Controle.Page.Session("notamsg_" & Controle.ID & "_" & Controle.Page.Session.SessionID)
                    If IsNothing(nota) Then
                        nota = Controle.Page.Session("notamsg_" & Controle.Page.Session.SessionID)
                    End If

                    Return nota
                End Get
                Set(ByVal value As Msg)
                    Controle.Page.Session("notamsg_" & Controle.ID & "_" & Controle.Page.Session.SessionID) = value
                    Controle.Page.Session("notamsg_" & Controle.Page.Session.SessionID) = value
                End Set
            End Property

            ''' <summary>
            ''' Registro de nota padrão.
            ''' </summary>
            ''' <param name="Controle">Controle onde ocorreu o evento.</param>
            ''' <param name="Texto">Texto caso não seja um erro.</param>
            ''' <param name="Nota">Exception caso seja um erro.</param>
            ''' <param name="Identif">Identificação, que é o título da mensagem.</param>
            ''' <param name="IgnoreNotas">Se deve ignorar as notas ou não (não apresentá-las na tela).</param>
            ''' <param name="URLRedirNota">URL de encaminhamento. Vazio para postback (mesma tela).</param>
            ''' <remarks></remarks>
            Private Shared Sub Registra(ByVal Controle As Control, ByVal Texto As String, ByVal Nota As Exception, ByVal Identif As String, ByVal IgnoreNotas As Boolean, ByVal URLRedirNota As String)
                If Not IgnoreNotas Then
                    If IsNothing(Nota) And NZ(Texto, "") = "" Then
                        Controle.Page.Session.Remove("notamsg_" & Controle.ID & "_" & Controle.Page.Session.SessionID)
                        Controle.Page.Session.Remove("notamsg_" & Controle.Page.Session.SessionID)
                    Else
                        Ex(Controle) = New Msg(Controle, Texto, Nota, Identif, URLRedirNota)
                        If Not IsNothing(Nota) Then
                            ErroLogReg(Nota, NotaMsg.MsgTexto(Texto, Nota))
                        End If
                    End If
                End If
            End Sub

            ''' <summary>
            ''' Deve ser utilizada no TRY...CATCH ex as EXCEPTION...END TRY na forma TRATA(PAGE, EX, IGNORENotaS, URLREDIRNota).
            ''' </summary>
            ''' <param name="Controle">Página corrente onde ocorreu o Nota.</param>
            ''' <param name="Nota">Exception ocorrido.</param>
            ''' <param name="IgnoreNotas">Indica que Notas devem ser ignorados.</param>
            ''' <param name="URLRedirNota">Endereço da URL que receberá o controle. Esta deverá possuir NotaMSG.VERIFICA para que apresente o Nota adequadamente. Ignorar este parâmetro ou colocar "" significará POSTBACK na própria página.</param>
            ''' <remarks></remarks>
            Public Shared Sub Registra(ByVal Controle As Control, ByVal Nota As Exception, ByVal Identif As String, Optional ByVal IgnoreNotas As Boolean = False, Optional ByVal URLRedirNota As String = "")
                Registra(Controle, "", Nota, Identif, IgnoreNotas, URLRedirNota)
            End Sub

            ''' <summary>
            ''' Registro de mensagem texto para aparecimento mediante VERIFICA.
            ''' </summary>
            ''' <param name="Controle">Controle onde ocorre a mensagem.</param>
            ''' <param name="Texto">Texto que deverá aparecer na mensagem.</param>
            ''' <param name="Identif">Identificação ou título da mensagem.</param>
            ''' <param name="IgnoreNotas">Ignorar as notas.</param>
            ''' <param name="URLRedirNota">URL de redirecionamento. Ocorrerá o encaminhamento para esta tela onde a função VERIFICA deverá ser executada.</param>
            ''' <remarks></remarks>
            Public Shared Sub Registra(ByVal Controle As Control, ByVal Texto As String, ByVal Identif As String, Optional ByVal IgnoreNotas As Boolean = False, Optional ByVal URLRedirNota As String = "")
                Registra(Controle, Texto, Nothing, Identif, IgnoreNotas, URLRedirNota)
            End Sub

            ''' <summary>
            ''' Retorna a mensagem a ser exibida na tela a partir do texto simplificado e/ou Exception.
            ''' </summary>
            ''' <param name="Texto">Texto simplificado a ser incluído na explicação.</param>
            ''' <param name="Ex">Exception para tratamento através de messageEx.</param>
            ''' <returns>Retorna o texto representativo da notificação.</returns>
            ''' <remarks></remarks>
            Public Shared Function MsgTexto(ByVal Texto As String, ByVal Ex As Exception) As String
                Dim MsgNota As String = Texto
                If Not IsNothing(Ex) Then
                    MsgNota &= IIf(MsgNota <> "", vbCrLf, "") & MessageEx(Ex)
                End If
                Return MsgNota
            End Function

            ''' <summary>
            ''' Deve ser colocado em PRE_RENDER (postback ou não) das páginas que tratarão dos Notas. Serão apresentados os diálogos sem a necessidade de novo POST.
            ''' </summary>
            ''' <param name="Controle">Página atual sendo a propriedade PAGE, utilizada para recuperar variáveis de sessão.</param>
            ''' <remarks></remarks>
            Public Shared Sub Verifica(ByVal Controle As Control)
                Dim erro As Boolean = False
                Dim Nota As Msg = Ex(Controle)
                If Not IsNothing(Nota) Then
                    ' se existir erro, vermelho, caso contrário, azul
                    If Not IsNothing(Nota.Ex) Then
                        erro = True
                    End If

                    ' nota é texto mais erro em outra linha
                    Dim NotaMsg As String = MsgTexto(Nota.Texto, Nota.Ex)

                    ' apenas para compatibilidade com versão anterior
                    Dim ctl As Object = Nothing
                    If TypeOf Controle Is Page Then
                        ctl = Form.FindControl(Controle, "IcraftNotaMsg_Alert")
                    Else
                        ctl = Controle
                    End If

                    Dim add As Boolean = False
                    If IsNothing(ctl) Then
                        ctl = Controle.Page.LoadControl("~\uc\icftmessage\icftmessage.ascx")
                        add = True
                    End If


                    Try
                        ctl.NotaIdentif = Nota.Identif
                        ctl.NotaMsg = NotaMsg
                        ctl.Attributes("Icone") = IIf(erro, "erro", "info")
                        ctl.attributes("Escondido") = False
                        ctl.attributes("Botoes") = "OK"
                        ctl.attributes("EventosStr") = ""

                        If Nota.URLRedirNota.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase) Then
                            ctl.attributes("EventosStr") = Nota.URLRedirNota & ";" & "return false;"
                        Else
                            ctl.attributes("EventosStr") = "javascript:" & ctl.Fecha()
                            If Nota.URLRedirNota <> "" Then
                                ctl.attributes("EventosStr") &= "window.location='" & HttpContext.Current.Server.UrlEncode(Nota.URLRedirNota) & "';return false;"
                            End If
                        End If


                        ctl.databind()


                        If add Then
                            Controle.Page.Form.Controls.Add(ctl)
                        End If

#If _MYTYPE = "Web" Then
                        Dim mpe As Object = Form.FindControl(ctl, "mpeMsg")
                        mpe.Show()
#End If

                    Catch ex As Exception
                        ShowJSMessage(Controle.Page, "Nota: " & NotaMsg)
                    End Try
                    Limpa(Controle)
                End If
            End Sub

            Public Shared Sub Limpa(ByVal Controle As Control)
                If Not IsNothing(Ex(Controle)) Then
                    Ex(Controle) = Nothing
                End If
            End Sub


            Public Sub New()

            End Sub
        End Class



        Public Class Info
            Public Shared Function verificaMetodos(ByVal obj As Object) As String()
                Dim X() As String = {"ME", "MI", "COMIGO", "O", "A", "LHE"}
                System.Console.WriteLine(From PALAV As String In X Select PALAV Order By PALAV)

                Return (From mInfo As Reflection.MethodInfo In obj.GetType.GetMethods() Select mInfo.Name & "(" & String.Join(", ", (From pInfo As Reflection.ParameterInfo In mInfo.GetParameters() Select pInfo.Name).ToArray) & ")" & IIf(mInfo.ReturnType.Name = "Void", "", " As " & mInfo.ReturnType.Name).ToString).ToArray
            End Function

            Public Shared Function verificaPropriedades(ByVal obj As Object) As String()
                Return (From pInfo As Reflection.PropertyInfo In obj.GetType.GetProperties Select IIf(pInfo.CanRead, "", "WriteOnly ").ToString & IIf(pInfo.CanWrite, "", "ReadOnly ").ToString & pInfo.Name & "() As " & pInfo.PropertyType.Name).ToArray
            End Function
        End Class



        Class Email
            Private _completo As String = ""
            Private _soendereco As String = ""
            Private _descricao As String = ""
            Private _dominio As String = ""
            Shared ReadOnly Property Valida(ByVal Email As String) As Boolean
                Get
                    Return Regex.IsMatch(EmailStr(Email), "(^|[ \t\[\<\>\""]*)([\w-.]+@[\w-]+(\.[\w-]+)+)(($|[ \t\<\>\""]*))")
                End Get
            End Property
            Public ReadOnly Property Valida() As Boolean
                Get
                    Return Valida(_completo)
                End Get
            End Property
            Sub New(ByVal Email As String)
                _completo = EmailStr(Email)
                _soendereco = SoEmailStr(_completo)
                _descricao = RegexGroup(_completo, "\""(.*)\""", 1).Value
                _dominio = RegexGroup(_soendereco, "@(.*)$", 1).Value
            End Sub
            Public ReadOnly Property Dominio() As String
                Get
                    Return _dominio
                End Get
            End Property
            Public ReadOnly Property Completo() As String
                Get
                    Return _completo
                End Get
            End Property
            Public ReadOnly Property SoEndereco() As String
                Get
                    Return _soendereco
                End Get
            End Property
            Public ReadOnly Property Descricao() As String
                Get
                    Return _descricao
                End Get
            End Property
        End Class




        Public Shared Function ImagemPath(ByVal Page As Page, ByVal Arquivo As String) As String
            If Arquivo.StartsWith("~/") Then
                Return Page.ResolveUrl(Arquivo)
            ElseIf Arquivo.StartsWith("http://", StringComparison.OrdinalIgnoreCase) Then
                Return Arquivo
            End If
            Return Page.ResolveUrl("~/img/" & Arquivo)
        End Function

        Public Shared ReadOnly Property ThemePath(ByVal page As Page, ByVal Arquivo As String, Optional ByVal Theme As String = "Default", Optional ByVal Path As String = "~/App_Theme") As String
            Get
                Return page.ResolveUrl(URLExpr(Path, Theme, Arquivo))
            End Get
        End Property

        Shared Function StrToByteArray(ByVal Texto As String) As Byte()
            Try
                Dim Cod As New System.Text.ASCIIEncoding
                Return Cod.GetBytes(Texto)
            Catch
                Return Nothing
            End Try
        End Function

        Shared Function ByteArrayToStr(ByVal Bytes() As Byte) As String
            Try
                Dim Cod As New System.Text.ASCIIEncoding
                Return Cod.GetString(Bytes)
            Catch
                Return Nothing
            End Try
        End Function

        Shared Function TiraAcento(ByVal Texto As String) As String
            Dim de As String = "ÁÉÍÓÚáéíóúÇçÀàÃãÕõÂâÊêÈèÔô"
            Dim para As String = "AEIOUaeiouCcAaAaOoAaEeEeOo"
            For z = 1 To Len(de)
                Texto = Texto.Replace(Mid(de, z, 1), Mid(para, z, 1))
            Next
            Return Texto
        End Function

        Shared Function ObtemPag(ByVal URL As String, Optional ByVal Encoding As System.Text.Encoding = Nothing) As String
            If IsNothing(Encoding) Then
                Encoding = New ASCIIEncoding
            End If
            Return ObtemPag(URL, "GET", Encoding)
        End Function

        Shared Function ObtemTexto(ByVal Origem As String) As String
            Dim Ret As String = ""
            Dim Arq As New System.IO.StreamReader(Origem)
            Ret = Arq.ReadToEnd
            Return Ret
        End Function

        Shared Function ObtemPag(ByVal URL As String, ByVal Metodo As String, ByVal Encoding As System.Text.Encoding, ByVal ParamArray ListaParams() As Object) As String
            Try
                Dim Req As HttpWebRequest = HttpWebRequest.Create(New System.Uri(URL))
                Req.Headers.Add("SessionId", HttpContext.Current.Session.SessionID)
                Dim Params As ArrayList = ParamArrayToArrayList(ListaParams)

                Dim PostP As New StringBuilder
                For z = 0 To Params.Count - 1 Step 2
                    PostP.Append(IIf(PostP.Length <> 0, "&", "") & Params(z) & "=")
                    PostP.Append(Params(z + 1))
                Next

                ' trata parâmetros
                If Compare(Metodo, "POST") Then
                    Dim PostPBytes() As Byte = Encoding.GetBytes(PostP.ToString)
                    Req.Method = "POST"
                    Req.ContentType = "application/x-www-form-urlencoded"
                    Req.ContentLength = PostPBytes.Length
                    Req.GetRequestStream().Write(PostPBytes, 0, PostPBytes.Length)
                End If

                With New StreamReader(Req.GetResponse().GetResponseStream(), Encoding)
                    Return .ReadToEnd()
                End With
            Catch EX As Exception
                Return Icraft.IcftBase.MessageEx(EX, "Obtendo página")
            End Try
        End Function






        ''' <summary>
        ''' Capitaliza um string levando em consideração exceções como sigla de Estados.
        ''' </summary>
        ''' <param name="value">O valor a ser capitalizado.</param>
        ''' <returns>Retorna a string capitalizada.</returns>
        ''' <remarks></remarks>
        Public Shared Function PrimLetraMaius(ByVal value As String) As String
            Dim maiusc As String() = "AC;AL;AM;AP;BA;CE;DF;ES;GO;MA;MG;MS;MT;PA;PB;PE;PI;PR;RJ;RN;RO;RR;RS;SC;SE;SP;TO;IASERJ;HSE;UNIRIO;UERJ;UFRJ;UFF;UFSP;SBD;ABC;UFJF;INAMPS;SUS;UNICAMP;UFS;CEDER;UNOESTE;PUC;UFES;FMUSP;UFRGS;UFPR;FMJ;HUSC;UFP;HUT;CEDEM;HUCFF;EPM;HC;FFFCM;UNIFESP".Split(";")
            Dim minusc As String() = "da;de;di;do;du;das;des;dis;dos;dus;na;nas;no;nos;em;aos;ao".Split(";")

            Dim tempValue As Collections.IEnumerator = Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(LCase(value.Replace("'", " "))).Split(" ", ".", "/", "-").GetEnumerator

            Dim gen As New System.Text.StringBuilder()

            While tempValue.MoveNext
                If maiusc.Contains(UCase(tempValue.Current)) Then
                    gen.Append(UCase(tempValue.Current))
                ElseIf minusc.Contains(LCase(tempValue.Current)) Then
                    gen.Append(LCase(tempValue.Current))
                Else
                    gen.Append(tempValue.Current)
                End If

                If gen.Length < value.Length Then
                    gen.Append(value.Chars(gen.Length))
                End If
            End While

            Return gen.ToString()
        End Function

        ''' <summary>
        ''' Formas possíveis de entificar um texto.
        ''' </summary>
        ''' <remarks></remarks>
        Enum TipoEntifica
            Tudo
            SoAcentos
            MenosHTML
        End Enum

        ''' <summary>
        ''' Transforma caracteres especiais em seus respectivos códigos Asc, preparados para HTML.
        ''' </summary>
        ''' <param name="Texto">O texto que será entificado.</param>
        ''' <param name="Tipo">O tipo de entificação.</param>
        ''' <returns>Retorna o texto passado com caracteres preparados para HTML de acordo com Tipo.</returns>
        ''' <remarks></remarks>
        Shared Function Entifica(ByVal Texto As String, Optional ByVal Tipo As TipoEntifica = TipoEntifica.Tudo) As String
            Dim G1() As String = {"&", "¡", "¢", "£", "¤", "¥", "¦", "§", "¨", "©", "ª", "«", "¬", "¬", "®", "¯", "°", "±", "²", "³", "´", "µ", "¶", "•", "¸", "¹", "º", "»", "¼", "½", "¾", "¿", "×", "÷", "Æ", "Ð", "Ø", "Þ", "ß", "æ", "ø", "þ"}
            Dim G2() As String = {"À", "Á", "Â", "Ã", "Ä", "Å", "Ç", "È", "É", "Ê", "Ë", "Ì", "Í", "Î", "Ï", "Ñ", "Ò", "Ó", "Ô", "Õ", "Ö", "Ù", "Ú", "Û", "Ü", "Ý", "à", "á", "â", "ã", "ä", "å", "ç", "è", "é", "ê", "ë", "ì", "í", "î", "ï", "ð", "ñ", "ò", "ó", "ô", "õ", "ö", "ù", "ú", "û", "ü", "ý", "ÿ"}
            Dim G3() As String = {"""", "'", "<", ">"}

            If Tipo = TipoEntifica.Tudo Then
                For Each IT As String In G1
                    Texto = Texto.Replace(IT, "&#" & Asc(IT) & ";")
                Next
            End If

            If Tipo = TipoEntifica.Tudo Or Tipo = TipoEntifica.SoAcentos Then
                For Each IT As String In G2
                    Texto = Texto.Replace(IT, "&#" & Asc(IT) & ";")
                Next
            End If

            If Tipo = TipoEntifica.Tudo And Not Tipo = TipoEntifica.MenosHTML Then
                For Each IT As String In G3
                    Texto = Texto.Replace(IT, "&#" & Asc(IT) & ";")
                Next
            End If

            Return Texto
        End Function

        Shared Sub DebugPrint(ByVal page As Page, ByVal Conteudo As String)
            Dim Ctl As Label = page.Master.FindControl("txtDebug")
            If Not IsNothing(Ctl) Then
                Ctl.Text &= HttpContext.Current.Server.HtmlEncode(Conteudo) & "<br/>"
            End If
        End Sub

        Shared Sub AtribQuandoDef(ByRef VarByRef As Object, ByVal Def As Object)
            If Not IsNothing(Def) Then
                VarByRef = Def
            End If
        End Sub

        ''' <summary>
        ''' Classe responsável por criar objetos com base em uma dll carregada dinamicamente.
        ''' </summary>
        ''' <remarks></remarks>
        Public Class CriadorDeObjetos
            Private _assembly As Reflection.Assembly

            Public Sub New(ByVal dll As String)
                Dim dirList As String = ""
                Dim errList As String = ""

#If _MyType = "Web" Then
                dirList = HttpContext.Current.Server.MapPath("~/bin/" & dll) & ";" & HttpContext.Current.Server.MapPath(dll) & ";"
#ElseIf _MyType = "WindowsForms" Then
                dirList = Application.StartupPath & "\" & dll & ";" & Application.StartupPath & "\bin\" & dll & ";" & dll & ";"
#ElseIf _MyType = "Console" Then
                Dim init As String = System.AppDomain.CurrentDomain.BaseDirectory
                dirList = init & "\" & dll & ";" & init & "\bin\" & dll & ";" & dll & ";"
#End If

                errList = dirList
                For Each s As String In dirList.Split(";")
                    If IO.File.Exists(s) Then
                        _assembly = Reflection.Assembly.LoadFile(s)
                        Exit For
                    Else
                        errList = errList.Replace(s & ";", "")
                    End If
                Next

                If String.IsNullOrEmpty(errList) Then
                    Throw New Exception("A dll especificada não foi localizada")
                End If

                If _assembly Is Nothing Then
                    Throw New Exception("Não foi possível carregar a dll especificada. Verifique se a mesma contém um formato válido.")
                End If
            End Sub

            ''' <summary>
            ''' Cria um objeto de um tipo especificado dentro da dll com base em seu nome.
            ''' </summary>
            ''' <param name="obj">O nome do tipo do objeto que será criado.</param>
            ''' <param name="params">Parâmetros que atendam a algum construtor do objeto especificado. No caso de nada ser passado, o contrutor default será admitido.</param>
            ''' <returns>Retorna o objeto criado.</returns>
            ''' <remarks></remarks>
            Public Function Criar(ByVal obj As String, ByVal ParamArray params() As Object) As Object
                Return _assembly.CreateInstance(getTipo(obj).FullName, False, Reflection.BindingFlags.CreateInstance, Nothing, params, Nothing, Nothing)
            End Function

            Public Shared Function Criar(ByVal assemblyInstance As Reflection.Assembly, ByVal objectName As String, ByVal ParamArray params() As Object) As Object
                Return assemblyInstance.CreateInstance(getTipo(assemblyInstance, objectName).FullName, False, Reflection.BindingFlags.CreateInstance, Nothing, params, Nothing, Nothing)
            End Function

            Public Function getTipo(ByVal typeName As String) As System.Type
                Return _assembly.GetTypes().Single(Function(tp) tp.Name = typeName)
            End Function

            Public Shared Function getTipo(ByVal asb As Reflection.Assembly, ByVal typeName As String) As System.Type
                Return asb.GetTypes().Single(Function(tp) tp.Name = typeName)
            End Function

            Public Function verificaTipos() As String()
                Return (From assemb As System.Type In _assembly.GetTypes Select assemb.Name).ToArray
            End Function

            Public Shared Function verificaTipos(ByVal asb As Reflection.Assembly) As String()
                Return (From assemb As System.Type In asb.GetTypes Select assemb.Name).ToArray
            End Function

            Public ReadOnly Property Assembly() As Reflection.Assembly
                Get
                    Return _assembly
                End Get
            End Property
        End Class

        Class PoolAsync
            Public MaxPends As Integer = Nothing
            Private _Pool As ArrayList = Nothing
            Sub New()
            End Sub
            Sub New(ByVal Pool As ArrayList, Optional ByVal MaxPends As Integer = 20)
                If Not IsNothing(Pool) Then
                    _Pool = Pool
                Else
                    _Pool = New ArrayList
                End If
                Me.MaxPends = NZ(MaxPends, 20)
            End Sub

            Private Function _Chama(ByVal Rotina As System.Delegate, Optional ByVal Param As Object = Nothing) As Threading.Thread
                Dim th As System.Threading.Thread = Nothing
                If Not IsNothing(Pend()) Then
                    If TypeOf (Rotina) Is System.Threading.ThreadStart Then
                        th = New System.Threading.Thread(CType(Rotina, System.Threading.ThreadStart))
                    ElseIf TypeOf (Rotina) Is System.Threading.ParameterizedThreadStart Then
                        th = New System.Threading.Thread(CType(Rotina, System.Threading.ParameterizedThreadStart))
                    End If
                    If Not IsNothing(_Pool) Then
                        _Pool.Add(th)
                    End If
                    If IsNothing(Param) Then
                        th.Start()
                    Else
                        th.Start(Param)
                    End If
                End If
                Return th
            End Function

            Function Chama(ByVal Rotina As System.Threading.ThreadStart) As Threading.Thread
                Return _Chama(Rotina, Nothing)
            End Function

            Function Chama(ByVal Rotina As System.Threading.ParameterizedThreadStart, ByVal Param As Object) As Threading.Thread
                Return _Chama(Rotina, Param)
            End Function

            Function Invoke(ByVal Sender As Object, ByVal Rotina As System.Delegate) As System.IAsyncResult
                Return _Invoke(Sender, Rotina)
            End Function

            Function Invoke(ByVal Sender As Object, ByVal Rotina As System.Delegate, ByVal Params As Object) As System.IAsyncResult
                Return _Invoke(Sender, Rotina, Params)
            End Function


            Function _Invoke(ByVal Sender As Object, ByVal Rotina As System.Delegate, Optional ByVal Param As Object = Nothing) As System.IAsyncResult
                Dim ar As Object = Nothing
                If Not IsNothing(Pend()) Then
                    If TypeOf (Rotina) Is System.Threading.ThreadStart Then
                        ar = Sender.BeginInvoke(Rotina)
                    ElseIf TypeOf (Rotina) Is System.Threading.ParameterizedThreadStart Then
                        ar = Sender.BeginInvoke(Rotina, Param)
                    End If
                    If Not IsNothing(_Pool) Then
                        _Pool.Add(ar)
                    End If
                End If
                Return ar
            End Function

            Public Function Pend() As Object
                If MaxPends = 0 Then
                    Return True
                End If
                If IsNothing(_Pool) Then
                    Return False
                End If
                Dim z As Integer = 0
                Do While z < _Pool.Count
                    Dim obj As Object = _Pool(z)
                    If TypeOf (obj) Is System.Threading.Thread AndAlso CType(obj, Threading.Thread).ThreadState = Threading.ThreadState.Stopped Then
                        _Pool.RemoveAt(z)
                    ElseIf obj.ToString = "System.Windows.Forms.Control+ThreadMethodEntry" AndAlso obj.iscompleted Then
                        _Pool.RemoveAt(z)
                    Else
                        z += 1
                    End If
                Loop
                If _Pool.Count >= MaxPends Then
                    Return Nothing
                End If
                Return z > 0
            End Function
        End Class


        ''' <summary>
        ''' Monta variável para armazenamento padronizado na sessão, viewstate ou condição correspondente a página em questão.
        ''' </summary>
        ''' <param name="Page">Página onde será obtido request para acesso à url da página.</param>
        ''' <param name="Bloco">Bloco ou condição específica na página. Utilizado para diferenciar controles, mas pode ser deixado em branco.</param>
        ''' <param name="Param">Nome da variável a ser registrada</param>
        ''' <value>String contendo variável de registro.</value>
        ''' <returns>String contendo variável de registro.</returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property VarDeSessao(ByVal Page As Page, ByVal Bloco As String, ByVal Param As String) As String
            Get
                Return UCase(Page.Request.Url.AbsolutePath & IIf(Bloco <> "", "_" & Bloco, "") & IIf(Param <> "", "_" & Param, ""))
            End Get
        End Property


        Public Shared Property Atrib(ByVal Objeto As Object, ByVal Atributo As String) As Object
            Get
                Try
                    Dim r As Reflection.MemberInfo = Objeto.GetType.GetMember(Atributo, Reflection.BindingFlags.IgnoreCase Or Reflection.BindingFlags.Public Or Reflection.BindingFlags.Instance)(0)
                    If r.MemberType = Reflection.MemberTypes.Property Then
                        Return Objeto.GetType.GetProperty(r.Name).GetValue(Objeto, Nothing, Nothing, Nothing, Nothing)
                    ElseIf r.MemberType = Reflection.MemberTypes.Field Then
                        Return Objeto.GetType.GetField(r.Name).GetValue(Objeto)
                    Else
                        Throw New Exception("Especificado tipo " & r.MemberType.ToString & " não tratado no objeto " & Objeto.GetType.ToString)
                    End If
                Catch
                    Throw New Exception("Objeto " & Objeto.GetType.ToString & " não possui membro " & Atributo)
                End Try
                Return Nothing
            End Get
            Set(ByVal value As Object)
                Try
                    Dim r As Reflection.MemberInfo = Objeto.GetType.GetMember(Atributo, Reflection.BindingFlags.IgnoreCase Or Reflection.BindingFlags.Public Or Reflection.BindingFlags.Instance)(0)
                    If r.MemberType = Reflection.MemberTypes.Property Then
                        Dim p As System.Reflection.PropertyInfo = Objeto.GetType.GetProperty(r.Name)
                        p.SetValue(Objeto, Convert.ChangeType(value, p.PropertyType), Reflection.BindingFlags.SetProperty, Nothing, Nothing, Nothing)
                    ElseIf r.MemberType = Reflection.MemberTypes.Field Then
                        Dim m As System.Reflection.FieldInfo = Objeto.GetType.GetField(r.Name)
                        m.SetValue(Objeto, Convert.ChangeType(value, m.FieldType), Reflection.BindingFlags.SetField, Nothing, Nothing)
                    Else
                        Throw New Exception("Especificado tipo " & r.MemberType.ToString & " não tratado no objeto " & Objeto.GetType.ToString)
                    End If
                Catch
                    Throw New Exception("Objeto " & Objeto.GetType.ToString & " não possui membro " & Atributo)
                End Try

            End Set
        End Property


        ''' <summary>
        ''' Pega titulo da página entre as tags title.
        ''' </summary>
        ''' <param name="Arquivo">Nome do arquivo de página (pode conter ~/).</param>
        ''' <returns>Retorna o texto entre as tags title da página.</returns>
        ''' <remarks></remarks>
        Shared Function PegaTituloPagWeb(ByVal Arquivo As String) As String
            Return PegaHtmlEmArquivo(FileExpr(Arquivo), "title").Inner
        End Function


        ''' <summary>
        ''' Pegar título da página web a partir do cabeçalho
        ''' </summary>
        ''' <param name="Arquivo">Arquivo de página web (pode conter ~/).</param>
        ''' <returns>Retorna trecho de título do cabeçalho.</returns>
        ''' <remarks></remarks>
        Shared Function PegaTituloPagWebDoCabeca(ByVal Arquivo As String) As String
            Dim Arq As New System.IO.StreamReader(FileExpr(Arquivo), System.Text.Encoding.Default)
            Dim Txt As String = Arq.ReadToEnd
            Arq.Close()
            Dim Txt2 As String = Icraft.IcftBase.RegexGroup(Txt, "(?is)title=""(.*?)""", 1).Value
            Return Txt2
        End Function



        ''' <summary>
        ''' Pega um regexhtml de uma página.
        ''' </summary>
        ''' <param name="Arquivo">Arquivo a ser pesquisado (pode conter ~/).</param>
        ''' <param name="Tag">Tag que será pesquisada.</param>
        ''' <returns>Retorna regexhtml daquele trecho.</returns>
        ''' <remarks></remarks>
        Shared Function PegaHtmlEmArquivo(ByVal Arquivo As String, ByVal Tag As String) As Icraft.IcftBase.RegexHtml
            Dim Arq As New System.IO.StreamReader(FileExpr(Arquivo), System.Text.Encoding.Default)
            Dim Txt As String = Arq.ReadToEnd
            Arq.Close()
            Return New Icraft.IcftBase.RegexHtml(Txt, Tag)
        End Function

        Public Shared Function NovaSenha(Optional ByVal Qtd As Integer = 6, Optional ByVal Maiusc As Boolean = True, Optional ByVal Minusc As Boolean = False, Optional ByVal Complex As Boolean = False, Optional ByVal Numeros As Boolean = False) As String
            Dim ComplexStr As String = "!@#$&*_-+=?|\/"
            Dim NR As Integer = 0
            If Maiusc Then
                NR += 26
            End If
            If Minusc Then
                NR += 26
            End If
            If Numeros Then
                NR += 10
            End If
            If Complex Then
                NR += Len(ComplexStr)
            End If

            Dim Senha As String = ""
            For Z = 1 To Qtd
                Dim Carac As Integer = Int(Rnd(Rnd() * 100) * NR)
                If Maiusc And Carac < 26 Then ' A-Z
                    Senha += Chr(Carac + Asc("A"))
                ElseIf Minusc And Carac < IIf(Maiusc, 26, 0) + 26 Then ' a-z
                    Senha += Chr(Carac - IIf(Maiusc, 26, 0) + Asc("a"))
                ElseIf Numeros And Carac < IIf(Maiusc, 26, 0) + IIf(Minusc, 26, 0) + 10 Then ' 0-9
                    Senha += Chr(Carac - IIf(Maiusc, 26, 0) - IIf(Minusc, 26, 0) + Asc("0"))
                ElseIf Complex And Carac >= IIf(Maiusc, 26, 0) + IIf(Minusc, 26, 0) + IIf(Numeros, 10, 0) Then ' complex
                    Senha += ComplexStr.Chars(Carac - IIf(Maiusc, 26, 0) - IIf(Minusc, 26, 0) - IIf(Numeros, 10, 0))
                End If
            Next
            Return Senha
        End Function



        Public Class Gerador
            ''' <summary>
            ''' Especificação de gravação de tipo em GERADOR para todas, parte um ou parte dois.
            ''' </summary>
            ''' <remarks></remarks>
            Enum GravaOracleParteTipo
                Todas
                Um
                Dois
            End Enum

            Sub New()
                Importa = Soma(System.Enum.GetValues(GetType(Criterios)))
            End Sub

            Public TabsSistema As String = "SYS_CONFIG_GLOBAL;SYS_CONFIG_USUARIO;SYS_DELETE;SYS_LOCALID;SYS_OCORRENCIA"
            Public TabsGerador As String = "GER_ADICIONAL_OBJ;GER_CAMPO;GER_CLASSE;GER_DIREITO;GER_GRUPO;GER_INDICE;GER_RELACIONAMENTO;GER_SISTEMA;GER_TABELA;GER_USUARIO;GER_VISAO"




            Public Sistema As String
            Public Descr As String
            Public Ver As String
            Public Data_Ver As Date
            Public Rev As String
            Public ListaDeTabelas As List(Of Tabela)
            Public ListaDeRels As List(Of Rel)
            Public ListaDeClasses As List(Of Classe)
            Public ListaDeVisoes As List(Of Visao)
            Public ListaDeObjs As List(Of Obj)
            Public ListaDeUsuarios As List(Of Usuario)
            Public ListaDeDireitos As List(Of Direito)
            Public Tipo As TipoBaseSQL
            Public Importa As Criterios
            Public Exporta As Criterios

            <Flags()> _
            Public Enum Criterios
                InfraSistema = 1
                InfraGerador = 2
                DadosGerador = 4
                Iniciar = 8
                Incluir = 16
                Sistema = 32
                Classe = 64
                Tabela = 128
                Campo = 256
                Comentario = 512
                Relacionamento = 1024
                Visao = 2048
                Objeto = 4096
                Usuario = 8192
                Direito = 16384
            End Enum


            Public ReadOnly Property ListaDeCampos() As List(Of Campo)
                Get
                    Dim Campos As New List(Of Campo)
                    For Each TAB As Tabela In ListaDeTabelas
                        Campos.AddRange(TAB.ListaDeCampos)
                    Next
                    Return Campos
                End Get
            End Property

            Class Classe
                Public Overrides Function tostring() As String
                    Return ""
                End Function

                Public Classe As String
                Public Descr As String
                Sub New(ByVal Classe As String, ByVal Descr As String)
                    Me.Classe = Classe
                    Me.Descr = Descr
                End Sub
                Default Public Property Attributes(ByVal Nome As String) As Object
                    Get
                        If Compare(Nome, "Classe") Then
                            Return Me.Classe
                        ElseIf Compare(Nome, "Descr") Then
                            Return Me.Descr
                        Else
                            Throw New Exception("Propriedade inválida " & Nome & " para objeto " & Me.GetType.ToString & ".")
                        End If
                    End Get
                    Set(ByVal value As Object)
                        If Compare(Nome, "Classe") Then
                            Me.Classe = value
                        ElseIf Compare(Nome, "Descr") Then
                            Me.Descr = value
                        Else
                            Throw New Exception("Propriedade inválida " & Nome & " para objeto " & Me.GetType.ToString & ".")
                        End If
                    End Set
                End Property
            End Class
            Class Visao
                Public Nome As String
                Public Classe As String
                Public Texto As String
                Sub New(ByVal Nome As String, ByVal Classe As String, ByVal Texto As String)
                    Me.Nome = Nome
                    Me.Classe = Classe
                    Me.Texto = Texto
                End Sub
            End Class
            Class Obj
                Public Tipo As String
                Public Ordem As Integer
                Public Texto As String
                Public Descr As String
                Sub New(ByVal Tipo As String, ByVal Ordem As Integer, ByVal Texto As String, ByVal Descr As String)
                    Me.Tipo = Tipo
                    Me.Ordem = Ordem
                    Me.Texto = Texto
                    Me.Descr = Descr
                End Sub
            End Class
            Class Usuario
                Public Login As String
                Public Grupo As String
                Public Senha As String
                Public Nome As String
                Public Depto As String
                Public Obs As String

                Public Function ListaDeDireitos() As List(Of Direito)
                    Dim Ret As New List(Of Direito)
                    For Each Dr As Direito In Me.ListaDeDireitos
                        If Dr.Usuario = Login Then
                            Ret.Add(Dr)
                        End If
                    Next
                    Return Ret
                End Function

                Sub New(ByVal Login As String, ByVal Grupo As String, ByVal Senha As String, ByVal Nome As String, ByVal Depto As String, ByVal Obs As String)
                    Me.Login = Login
                    Me.Grupo = Grupo
                    Me.Senha = Senha
                    Me.Nome = Nome
                    Me.Depto = Depto
                    Me.Obs = Obs
                End Sub
            End Class
            Class Direito
                Public Tipo As String
                Public Objeto As String
                Public Usuario As String
                Public Permissao As String
                Sub New(ByVal Tipo As String, ByVal Objeto As String, ByVal Usuario As String, ByVal Permissao As String)
                    Me.Tipo = Tipo
                    Me.Objeto = Objeto
                    Me.Usuario = Usuario
                    Me.Permissao = Permissao
                End Sub
            End Class
            Class Tabela
                Public Overrides Function tostring() As String
                    Return Tabela
                End Function

                Public Tabela As String
                Public Ordem As Integer
                Public Chave_Prima As String
                Public Codigo As String
                Public ListaDeCampos As List(Of Campo)
                Public Classe As String
                Public Descr As String
                Public Etiq As String
                Public Function Diferenca(ByVal Tab As Tabela, Optional ByVal page As Page = Nothing) As String
                    Dim Result As String = ""

                    If Prop("chktab12", "checked", page) Then
                        If Me.Tabela <> Tab.Tabela Then
                            Result &= "Tabela: " & NZV(Me.Tabela, "<vazio>") & ComboSepDefault & NZV(Tab.Tabela, "<vazio>") & vbCrLf
                            Return Result
                        End If
                        If Me.Classe <> Tab.Classe Then
                            Result &= "Classe: " & NZV(Me.Classe, "<vazio>") & ComboSepDefault & NZV(Tab.Classe, "<vazio>") & vbCrLf
                        End If
                        If Me.Chave_Prima <> Tab.Chave_Prima Then
                            Result &= "Chave Primária: " & NZV(Me.Chave_Prima, "<vazio>") & ComboSepDefault & NZV(Tab.Chave_Prima, "<vazio>") & vbCrLf
                        End If
                        If Me.Descr <> Tab.Descr Then
                            Result &= "Descrição em Português: " & NZV(Me.Descr, "<vazio>") & ComboSepDefault & NZV(Tab.Descr, "<vazio>") & vbCrLf
                        End If
                        If Me.Etiq <> Tab.Etiq Then
                            Result &= "Etiqueta em Português: " & NZV(Me.Etiq, "<vazio>") & ComboSepDefault & NZV(Tab.Etiq, "<vazio>") & vbCrLf
                        End If
                    End If

                    Dim Segm As String = ""
                    ' campos que existem em x e não existem e y

                    If Prop("chktab12camp1s2", "checked", page) Then
                        Dim NomeCampos As ArrayList = ItemsToArrayList(Tab.ListaDeCampos, "Campo")
                        For Each CampoX As Campo In Me.ListaDeCampos
                            If NomeCampos.IndexOf(CampoX.Campo) = -1 Then
                                Segm &= CampoX.Campo & ComboSepDefault & "<inexistente>" & vbCrLf
                            End If
                        Next
                    End If

                    ' campos em y sem existirem em x 

                    If Prop("chktab12camp2s1", "checked", page) Then
                        Dim NomeCampos As ArrayList = ItemsToArrayList(Me.ListaDeCampos, "Campo")
                        For Each CampoY As Campo In Tab.ListaDeCampos
                            If NomeCampos.IndexOf(CampoY.Campo) = -1 Then
                                Segm &= "<inexistente>" & ComboSepDefault & CampoY.Campo & vbCrLf
                            End If
                        Next
                    End If

                    ' existe nos dois e é diferente
                    If Prop("chktab12camp12", "checked", page) Then
                        Dim NomeCampos As ArrayList = ItemsToArrayList(Me.ListaDeCampos, "Campo")
                        For Each CampoY As Campo In Tab.ListaDeCampos
                            If NomeCampos.IndexOf(CampoY.Campo) <> -1 Then
                                Dim CampoX As Campo = Campos(CampoY.Campo)
                                Dim CampoDif As String = CampoX.Diferenca(CampoY)
                                If CampoDif <> "" Then
                                    Segm &= CampoDif & vbCrLf
                                End If
                            End If
                        Next
                    End If

                    If Segm <> "" Then
                        Result &= "Campos:" & vbCrLf & InsereTab(Segm, Gerador_Tabula)
                    End If
                    If Result <> "" Then
                        Return Me.Tabela & vbCrLf & InsereTab(Result, Gerador_Tabula)
                    End If
                    Return Result
                End Function
                Function Campos(ByVal Campo As String) As Object
                    Dim NomeCampos As ArrayList = ItemsToArrayList(Me.ListaDeCampos, "Campo")
                    Dim Pos As Integer = NomeCampos.IndexOf(Campo)
                    If Pos <> -1 Then
                        Return ListaDeCampos(Pos)
                    End If
                    Return Nothing
                End Function
                Public Sub New(ByVal Tabela As String, ByVal Ordem As Integer, ByVal Chave_Prima As String, ByVal Campos As List(Of Campo), ByVal Codigo As String, ByVal Classe As String, ByVal Descr As String, ByVal Etiq As String)
                    Me.Tabela = Tabela
                    Me.Ordem = Ordem
                    Me.Chave_Prima = Chave_Prima
                    Me.ListaDeCampos = Campos
                    Me.Codigo = Codigo
                    Me.Classe = Classe
                    Me.Descr = Descr
                    Me.Etiq = Etiq
                End Sub
                Default Public Property Attributes(ByVal Nome As String) As Object
                    Get
                        If Compare(Nome, "Tabela") Then
                            Return Me.Tabela
                        ElseIf Compare(Nome, "Ordem") Then
                            Return Me.Ordem
                        ElseIf Compare(Nome, "Chave_Prima") Then
                            Return Me.Chave_Prima
                        ElseIf Compare(Nome, "Codigo") Then
                            Return Me.Codigo
                        ElseIf Compare(Nome, "Campos") Then
                            Return Me.ListaDeCampos
                        ElseIf Compare(Nome, "Classe") Then
                            Return Me.Classe
                        ElseIf Compare(Nome, "Descr") Then
                            Return Me.Descr
                        ElseIf Compare(Nome, "Etiq") Then
                            Return Me.Etiq
                        Else
                            Throw New Exception("Propriedade inválida " & Nome & " para objeto " & Me.GetType.ToString & ".")
                        End If
                    End Get
                    Set(ByVal value As Object)
                        If Compare(Nome, "Tabela") Then
                            Me.Tabela = value
                        ElseIf Compare(Nome, "Ordem") Then
                            Me.Ordem = value
                        ElseIf Compare(Nome, "Chave_Prima") Then
                            Me.Chave_Prima = value
                        ElseIf Compare(Nome, "Codigo") Then
                            Me.Codigo = value
                        ElseIf Compare(Nome, "Campos") Then
                            Me.ListaDeCampos = value
                        ElseIf Compare(Nome, "Classe") Then
                            Me.Classe = value
                        ElseIf Compare(Nome, "Descr") Then
                            Me.Descr = value
                        ElseIf Compare(Nome, "Etiq") Then
                            Me.Etiq = value
                        Else
                            Throw New Exception("Propriedade inválida " & Nome & " para objeto " & Me.GetType.ToString & ".")
                        End If
                    End Set
                End Property
            End Class
            Class Campo
                Public Overrides Function tostring() As String
                    Return ""
                End Function
                Public Tabela As String = Nothing
                Public Campo As String
                Public Ordem As Integer
                Public Tipo_Access As String
                Public Tipo_Oracle As String
                Public Tipo_MySQL As String
                Public Etiq As String
                Public Descr As String
                Public Prop_Extend As String
                Public Function Diferenca(ByVal Campo As Campo) As String
                    Dim Result As String = ""
                    If Me.Campo <> Campo.Campo Then
                        Result &= Me.Campo & ComboSepDefault & Campo.Campo & vbCrLf
                        Return Result
                    End If

                    Dim Segm As String = ""
                    If Me.Tipo_Access <> Campo.Tipo_Access Then
                        Segm &= "Tipo Access: " & NZV(Me.Tipo_Access, "<vazio>") & ComboSepDefault & NZV(Campo.Tipo_Access, "<vazio>") & vbCrLf
                    End If
                    If Me.Tipo_Oracle <> Campo.Tipo_Oracle Then
                        Segm &= "Tipo Oracle: " & NZV(Me.Tipo_Oracle, "<vazio>") & ComboSepDefault & NZV(Campo.Tipo_Oracle, "<vazio>") & vbCrLf
                    End If
                    If Me.Tipo_MySQL <> Campo.Tipo_MySQL Then
                        Segm &= "Tipo MySQL: " & NZV(Me.Tipo_MySQL, "<vazio>") & ComboSepDefault & NZV(Campo.Tipo_MySQL, "<vazio>") & vbCrLf
                    End If
                    If Me.Etiq <> Campo.Etiq Then
                        Segm &= "Etiqueta em Português: " & NZV(Me.Etiq, "<vazio>") & ComboSepDefault & NZV(Campo.Etiq, "<vazio>") & vbCrLf
                    End If
                    If Me.Descr <> Campo.Descr Then
                        Segm &= "Descrição em Português: " & NZV(Me.Descr, "<vazio>") & ComboSepDefault & NZV(Campo.Descr, "<vazio>") & vbCrLf
                    End If
                    If Me.Prop_Extend <> Campo.Prop_Extend Then
                        Segm &= "Propriedades Extendidas: " & NZV(Me.Prop_Extend, "<vazio>") & ComboSepDefault & NZV(Campo.Prop_Extend, "<vazio>") & vbCrLf
                    End If
                    If Segm <> "" Then
                        Result &= Me.Campo & ":" & vbCrLf & InsereTab(Segm, Gerador_Tabula)
                    End If
                    Return Result
                End Function
                Public Sub New(ByVal Tabela As String, ByVal Campo As String, ByVal Ordem As Integer, ByVal Tipo_Access As String, ByVal Tipo_Oracle As String, ByVal Tipo_MySQL As String, ByVal Etiq As String, ByVal Descr As String, ByVal Prop_Extend As String)
                    Me.Tabela = Tabela
                    Me.Campo = Campo
                    Me.Ordem = Ordem
                    Me.Tipo_Access = UCase(NZ(Tipo_Access, ""))
                    Me.Tipo_Oracle = UCase(NZ(Tipo_Oracle, ""))
                    Me.Tipo_MySQL = UCase(NZ(Tipo_MySQL, ""))
                    Me.Etiq = Etiq
                    Me.Descr = Descr
                    Me.Prop_Extend = Prop_Extend
                End Sub
                Default Public Property Attributes(ByVal Nome As String) As Object
                    Get
                        If Compare(Nome, "Campo") Then
                            Return Me.Campo
                        ElseIf Compare(Nome, "Ordem") Then
                            Return Me.Ordem
                        ElseIf Compare(Nome, "Tipo_Access") Then
                            Return Me.Tipo_Access
                        ElseIf Compare(Nome, "Tipo_Oracle") Then
                            Return Me.Tipo_Oracle
                        ElseIf Compare(Nome, "Tipo_MySQL") Then
                            Return Me.Tipo_MySQL
                        ElseIf Compare(Nome, "Etiq") Then
                            Return Me.Etiq
                        ElseIf Compare(Nome, "Descr") Then
                            Return Me.Descr
                        ElseIf Compare(Nome, "Prop_Extend") Then
                            Return Me.Prop_Extend
                        Else
                            Throw New Exception("Propriedade inválida " & Nome & " para objeto " & Me.GetType.ToString & ".")
                        End If
                    End Get
                    Set(ByVal value As Object)
                        If Compare(Nome, "Campo") Then
                            Me.Campo = value
                        ElseIf Compare(Nome, "Ordem") Then
                            Me.Ordem = value
                        ElseIf Compare(Nome, "Tipo_Access") Then
                            Me.Tipo_Access = value
                        ElseIf Compare(Nome, "Tipo_Oracle") Then
                            Me.Tipo_Oracle = value
                        ElseIf Compare(Nome, "Etiq_Prop") Then
                            Me.Etiq = value
                        ElseIf Compare(Nome, "Descr") Then
                            Me.Descr = value
                        ElseIf Compare(Nome, "Prop_Extend") Then
                            Me.Prop_Extend = value
                        Else
                            Throw New Exception("Propriedade inválida " & Nome & " para objeto " & Me.GetType.ToString & ".")
                        End If
                    End Set
                End Property
            End Class
            Class Rel
                Public Overrides Function tostring() As String
                    Return ""
                End Function
                Public Nome As String
                Public Tabela_1 As String
                Public Campo_1 As String
                Public Tabela_N As String
                Public Campo_N As String
                Public Delete_Cascade As String
                Public Update_Cascade As String
                Public Obrig As Boolean
                Public Sub New(ByVal Nome As String, ByVal Tabela_1 As String, ByVal Campo_1 As String, ByVal Tabela_N As String, ByVal Campo_N As String, ByVal Delete_Cascade As String, ByVal Update_Cascade As String, ByVal Obrig As Boolean)
                    Me.Nome = Nome
                    Me.Tabela_1 = Tabela_1
                    Me.Campo_1 = Campo_1
                    Me.Tabela_N = Tabela_N
                    Me.Campo_N = Campo_N
                    Me.Delete_Cascade = Delete_Cascade
                    Me.Update_Cascade = Update_Cascade
                    Me.Obrig = Obrig
                End Sub
                Default Public Property Attributes(ByVal Nome As String) As Object
                    Get
                        If Compare(Nome, "Nome") Then
                            Return Me.Nome
                        ElseIf Compare(Nome, "Tabela_1") Then
                            Return Me.Tabela_1
                        ElseIf Compare(Nome, "Campo_1") Then
                            Return Me.Campo_1
                        ElseIf Compare(Nome, "Tabela_N") Then
                            Return Me.Tabela_N
                        ElseIf Compare(Nome, "Campo_N") Then
                            Return Me.Campo_N
                        ElseIf Compare(Nome, "Delete_Cascade") Then
                            Return Me.Delete_Cascade
                        ElseIf Compare(Nome, "Update_Cascade") Then
                            Return Me.Update_Cascade
                        ElseIf Compare(Nome, "Obrig") Then
                            Return Me.Obrig
                        Else
                            Throw New Exception("Propriedade inválida " & Nome & " para objeto " & Me.GetType.ToString & ".")
                        End If
                    End Get
                    Set(ByVal value As Object)
                        If Compare(Nome, "Nome") Then
                            Me.Nome = value
                        ElseIf Compare(Nome, "Tabela_1") Then
                            Me.Tabela_1 = value
                        ElseIf Compare(Nome, "Campo_1") Then
                            Me.Campo_1 = value
                        ElseIf Compare(Nome, "Tabela_N") Then
                            Me.Tabela_N = value
                        ElseIf Compare(Nome, "Campo_N") Then
                            Me.Campo_N = value
                        ElseIf Compare(Nome, "Delete_Cascade") Then
                            Me.Delete_Cascade = value
                        ElseIf Compare(Nome, "Update_Cascade") Then
                            Me.Update_Cascade = value
                        ElseIf Compare(Nome, "Obrig") Then
                            Me.Obrig = value
                        Else
                            Throw New Exception("Propriedade inválida " & Nome & " para objeto " & Me.GetType.ToString & ".")
                        End If
                    End Set
                End Property
                Public Function Diferenca(ByVal Rel As Rel) As String
                    Dim Result As String = ""
                    If Me.Nome <> Rel.Nome Then
                        Result &= Me.Nome & ComboSepDefault & Rel.Nome & vbCrLf
                        Return Result
                    End If

                    Dim Segm As String = ""
                    If Me.Obrig <> Rel.Obrig Then
                        Segm &= "Obrig: " & NZV(Me.Obrig, "<vazio>") & ComboSepDefault & NZV(Rel.Obrig, "<vazio>") & vbCrLf
                    End If
                    If Me.Tabela_1 <> Rel.Tabela_1 Then
                        Segm &= "Tabela_1: " & NZV(Me.Tabela_1, "<vazio>") & ComboSepDefault & NZV(Rel.Tabela_1, "<vazio>") & vbCrLf
                    End If
                    If Me.Tabela_N <> Rel.Tabela_N Then
                        Segm &= "Tabela_N: " & NZV(Me.Tabela_N, "<vazio>") & ComboSepDefault & NZV(Rel.Tabela_N, "<vazio>") & vbCrLf
                    End If
                    If Me.Campo_1 <> Rel.Campo_1 Then
                        Segm &= "Campo_1: " & NZV(Me.Campo_1, "<vazio>") & ComboSepDefault & NZV(Rel.Campo_1, "<vazio>") & vbCrLf
                    End If
                    If Me.Campo_N <> Rel.Campo_N Then
                        Segm &= "Campo_N: " & NZV(Me.Campo_N, "<vazio>") & ComboSepDefault & NZV(Rel.Campo_N, "<vazio>") & vbCrLf
                    End If
                    If Me.Delete_Cascade <> Rel.Delete_Cascade Then
                        Segm &= "Delete_Cascade: " & NZV(Me.Delete_Cascade, "<vazio>") & ComboSepDefault & NZV(Rel.Delete_Cascade, "<vazio>") & vbCrLf
                    End If
                    If Me.Update_Cascade <> Rel.Update_Cascade Then
                        Segm &= "Update_Cascade: " & NZV(Me.Update_Cascade, "<vazio>") & ComboSepDefault & NZV(Rel.Update_Cascade, "<vazio>") & vbCrLf
                    End If
                    If Segm <> "" Then
                        Result &= Me.Nome & ":" & vbCrLf & InsereTab(Segm, Gerador_Tabula)
                    End If
                    Return Result
                End Function
            End Class

            Function Tabelas(ByVal Tabela As String) As Gerador.Tabela
                Dim NomeTabs As ArrayList = ItemsToArrayList(Me.ListaDeTabelas, "Tabela")
                Dim Pos As Integer = NomeTabs.IndexOf(Tabela)
                If Pos <> -1 Then
                    Return ListaDeTabelas(Pos)
                End If
                Return Nothing
            End Function
            Function Classes(ByVal Classe As String) As Gerador.Classe
                Dim NomeClasses As ArrayList = ItemsToArrayList(Me.ListaDeClasses, "Classe")
                Dim Pos As Integer = NomeClasses.IndexOf(Classe)
                If Pos <> -1 Then
                    Return ListaDeClasses(Pos)
                End If
                Return Nothing
            End Function
            Function Rels(ByVal Nome As String) As Gerador.Rel
                Dim NomeRels As ArrayList = ItemsToArrayList(Me.ListaDeRels, "Nome")
                Dim Pos As Integer = NomeRels.IndexOf(Nome)
                If Pos <> -1 Then
                    Return ListaDeRels(Pos)
                End If
                Return Nothing
            End Function
            Function Visoes(ByVal Nome As String) As Gerador.Visao
                Dim NomeVisoes As ArrayList = ItemsToArrayList(Me.ListaDeVisoes, "Nome")
                Dim Pos As Integer = NomeVisoes.IndexOf(Nome)
                If Pos <> -1 Then
                    Return ListaDeVisoes(Pos)
                End If
                Return Nothing
            End Function
            Function Objs(ByVal Nome As String) As Gerador.Obj
                Dim NomeObjs As ArrayList = ItemsToArrayList(Me.ListaDeObjs, "Nome")
                Dim Pos As Integer = NomeObjs.IndexOf(Nome)
                If Pos <> -1 Then
                    Return ListaDeObjs(Pos)
                End If
                Return Nothing
            End Function
            Function NomeRel(ByVal Tabela_1 As String, ByVal Tabela_N As String, ByVal ListaDeRels As List(Of Gerador.Rel)) As String
                Dim Cod1 As String = Tabelas(Tabela_1).Codigo
                Dim CodN As String = Tabelas(Tabela_N).Codigo
                Dim Monta As String = Cod1 & "_" & CodN & "_"
                Dim Seq As Integer = 1
                For Each R As Rel In ListaDeRels
                    If RegexGroup(R.Nome, Monta & "[0-9]{1,2}").Success Then
                        Dim Num As Integer = Microsoft.VisualBasic.Right(R.Nome, 2)
                        If Num >= Seq Then
                            Seq = Num + 1
                        End If
                    End If
                Next
                Return Monta & Format(Seq, "00")
            End Function

            Public Sub CarregaGerador(ByVal Sistema As String, Optional ByVal StrGerador As String = "strgerador")
                Me.Sistema = Sistema
                Me.Tipo = TipoBaseSQL.Gerador

                ' detalhes do sistema
                If Importa And Criterios.Sistema Then
                    For Each r As DataRow In DSCarrega("select nome, descr, ver, data_ver, rev from sistema where nome=:sistema", StrGerador, ":sistema", Sistema).Tables(0).Rows
                        Me.Descr = NZ(r("descr"), "")
                        Me.Ver = NZ(r("ver"), "")
                        Me.Data_Ver = NZV(r("data_ver"), Nothing)
                        Me.Rev = NZ(r("rev"), "")
                    Next
                End If

                ' carrega classes
                Dim Classe As New List(Of Classe)
                If Importa And Criterios.Classe Then
                    For Each r As DataRow In DSCarrega("select classe, descr from classe_OBJ where sistema=:sistema", StrGerador, ":sistema", Sistema).Tables(0).Rows
                        Classe.Add(New Classe(r("Classe"), NZ(r("Descr"), "")))
                    Next
                End If
                Me.ListaDeClasses = Classe


                ' carrega tabelas e campos
                Dim Tabs As New List(Of Tabela)
                If Importa And Criterios.Tabela Then
                    For Each r As DataRow In DSCarrega("Select tabela, ordem, chave_prima, codigo, classe, descr, etiq from GER_tabela where sistema=:sistema", StrGerador, ":sistema", Sistema).Tables(0).Rows
                        Dim Cpos As New List(Of Campo)
                        If Importa And Criterios.Campo Then
                            For Each rr As DataRow In DSCarrega("select campo, ordem, tipo_access, tipo_oracle, tipo_mysql, etiq, descr, prop_extend from GER_campo where sistema=:sistema and tabela = :tabela order by ordem", StrGerador, ":sistema", Sistema, ":tabela", r("Tabela")).Tables(0).Rows
                                Cpos.Add(New Campo(r("tabela"), rr("Campo"), rr("Ordem"), NZ(rr("Tipo_Access"), ""), NZ(rr("Tipo_Oracle"), ""), NZ(rr("Tipo_MySQL"), ""), NZ(rr("Etiq"), ""), NZ(rr("Descr"), ""), NZ(rr("Prop_Extend"), "")))
                            Next
                        End If
                        Tabs.Add(New Tabela(r("Tabela"), r("Ordem"), NZ(r("Chave_Prima"), ""), Cpos, r("CODIGO"), r("CLASSE"), NZ(r("DESCR"), ""), NZ(r("ETIQ"), "")))
                    Next
                End If
                Me.ListaDeTabelas = Tabs

                ' carrega relacionamentos
                Dim Rels As New List(Of Rel)
                If Importa And Criterios.Relacionamento Then
                    For Each r As DataRow In DSCarrega("Select nome, tabela_1, campo_1, tabela_n, campo_n, delete_cascade, update_cascade, obrig from GER_relacionamento where sistema=:sistema", StrGerador, ":sistema", Sistema).Tables(0).Rows
                        Rels.Add(New Rel(r("Nome"), r("Tabela_1"), r("Campo_1"), r("Tabela_N"), r("Campo_N"), NZ(r("Delete_Cascade"), ""), NZ(r("Update_Cascade"), ""), r("Obrig")))
                    Next
                    Me.ListaDeRels = Rels
                End If

                ' carrega visoes
                Dim visoes As New List(Of Visao)
                If Importa And Criterios.Visao Then
                    For Each r As DataRow In DSCarrega("select VISAO, CLASSE, TEXTO FROM VISAO WHERE SISTEMA = :sistema", StrGerador, ":sistema", Sistema).Tables(0).Rows
                        visoes.Add(New Visao(r("VISAO"), r("CLASSE"), r("TEXTO")))
                    Next
                End If
                Me.ListaDeVisoes = visoes

                ' carrega objetos
                Dim objs As New List(Of Obj)
                If Importa And Criterios.Objeto Then
                    For Each r As DataRow In DSCarrega("select ordem, tipo, texto, descr from adicional_obj where sistema=:sistema order by ordem", StrGerador, ":sistema", Sistema).Tables(0).Rows
                        objs.Add(New Obj(r("tipo"), r("ordem"), r("texto"), NZ(r("descr"), "")))
                    Next
                End If
                Me.ListaDeObjs = objs

                ' carrega usuários
                Dim usuarios As New List(Of Usuario)
                If Importa And Criterios.Usuario Then
                    For Each r As DataRow In DSCarrega("select usuario,grupo,senha,nome,depto,obs from usuario where sistema=:sistema order by usuario", StrGerador, ":sistema", Sistema).Tables(0).Rows
                        usuarios.Add(New Usuario(r("usuario"), NZ(r("grupo"), ""), NZ(r("senha"), ""), NZ(r("nome"), ""), NZ(r("depto"), ""), NZ(r("obs"), "")))
                    Next
                End If
                ListaDeUsuarios = usuarios

                ' carrega permissões
                Dim Direitos As New List(Of Direito)
                If Importa And Criterios.Direito Then
                    For Each r As DataRow In DSCarrega("select tipo,objeto,usuario,permissao from direito where sistema=:sistema order by usuario,objeto", StrGerador, ":sistema", Sistema).Tables(0).Rows
                        Direitos.Add(New Direito(r("tipo"), r("objeto"), r("usuario"), r("permissao")))
                    Next
                End If
                ListaDeDireitos = Direitos
            End Sub

            Class DescrConcat

                ' descrição etiqueta e grupo
                Public DescrConcatFormat As String = "(?<g1>.*?)(\|(?<g2>.*?)(\|(?<g3>.*)|$)|$)"


                Private _tabela As String = ""
                Private _campo As String = ""
                Private _grupo As String = ""
                Private _etiq As String = ""
                Private _descr As String = ""

                Public ReadOnly Property Texto() As String
                    Get
                        Dim Ar As New ArrayList
                        If NZ(_grupo, "") <> "" Then
                            Ar.Add(Grupo)
                        End If
                        If NZ(_descr, "") <> "" Then
                            Ar.Add(_descr)
                        End If
                        If NZ(_etiq, "") <> "" Then
                            Ar.Add(_etiq)
                        End If
                        Return Join(Ar.ToArray, " | ")
                    End Get
                End Property
                Sub New(ByVal Tabela As String, ByVal Campo As String, ByVal Grupo As String, ByVal Etiq As String, ByVal Descr As String)
                    _tabela = Tabela
                    _campo = Campo
                    _grupo = Grupo
                    _etiq = Etiq
                    _descr = Descr
                End Sub
                Public Property Grupo() As String
                    Get
                        Return _grupo
                    End Get
                    Set(ByVal value As String)
                        _grupo = value
                    End Set
                End Property
                Public Property Etiq() As String
                    Get
                        Return _etiq
                    End Get
                    Set(ByVal value As String)
                        _etiq = value
                    End Set
                End Property
                Public Property Descr() As String
                    Get
                        Return _descr
                    End Get
                    Set(ByVal value As String)
                        _descr = value
                    End Set
                End Property
                Sub New(ByVal Tabela As String, ByVal Campo As String, ByVal Texto As String)
                    _tabela = Tabela
                    _campo = Campo

                    _descr = ""
                    _etiq = ""
                    _grupo = ""

                    Dim Txt() As String = Split(Texto, "|")
                    Select Case Txt.Count
                        Case 1
                            _descr = Txt(0).Trim()
                        Case 2
                            _descr = Txt(0).Trim()
                            _etiq = Txt(1).Trim()
                        Case 3
                            _grupo = Txt(0).Trim()
                            _descr = Txt(1).Trim()
                            _etiq = Txt(2).Trim()
                    End Select

                    If NZ(_etiq, "").EndsWith(".") Then
                        Dim troca As String = _etiq
                        _etiq = _descr
                        _descr = troca
                    End If

                    If NZ(_etiq, "") = "" Or NZ(_etiq, "") = "Geral" Then
                        If NZ(_campo, "") <> "" Then
                            _etiq = _campo
                        Else
                            _etiq = Tabela
                        End If
                    End If

                    If NZ(_grupo, "") = "" And NZ(_campo, "") <> "" Then
                        _grupo = "Geral"
                    End If

                End Sub
            End Class

            Public Sub CarregaMSAccess(ByVal Sistema As String, ByVal ArquivoMDB As String)

                Me.Sistema = Sistema
                Me.Tipo = TipoBaseSQL.MSAccess
                ArquivoMDB = FileExpr(ArquivoMDB)

                ' abre banco
                ' abre conexão
                Dim de As Object
                Try
                    de = CreateObject("DAO.DBEngine.36")
                Catch ex As Exception
                    Throw New Exception("Necessária inclusão de referência para DBEngine.36.")
                End Try

                Dim db As Object = de(0).OpenDatabase(ArquivoMDB)

                GaranteExistenciaDasListas(Importa And Criterios.Iniciar)

                Try

                    ' tabelas
                    If Importa And Criterios.Tabela Then
                        Dim TabOrdem As Integer = 1
                        For Each tbd As Object In db.TableDefs
                            If Not (tbd.Name Like "MSys*" Or tbd.Name Like "~*" Or (TemNaLista(TabsSistema, tbd.name) And Not CType(Importa And Criterios.InfraSistema, Boolean)) Or (TemNaLista(TabsGerador, tbd.name) And Not CType(Importa And Criterios.InfraGerador, Boolean))) Then

                                ' chave prima
                                Dim Chave_Prima As String = ""
                                Try
                                    For Each fld As Object In tbd.Indexes("PRIMARYKEY").Fields
                                        Chave_Prima &= IIf(Chave_Prima <> "", ";", "") & fld.Name
                                    Next
                                Catch
                                End Try

                                ' antes dos campos, inclui linha de tabela
                                Dim Descr As DescrConcat = New DescrConcat(tbd.Name, Nothing, DaoPropTab(tbd, "description"))

                                If Importa And Criterios.Classe Then
                                    Dim Z As Integer
                                    For Z = 0 To ListaDeClasses.Count - 1
                                        If ListaDeClasses(Z).Classe = Descr.Grupo Then
                                            Exit For
                                        End If
                                    Next
                                    If Z >= ListaDeClasses.Count Then
                                        ListaDeClasses.Add(New Gerador.Classe(Descr.Grupo, Descr.Grupo))
                                    End If
                                End If


                                ListaDeTabelas.Add(New Tabela(tbd.Name, TabOrdem, Chave_Prima, New List(Of Gerador.Campo), tbd.Name, Descr.Grupo, "", ""))

                                ' campos
                                If Importa And Criterios.Campo Then
                                    Dim CampoOrdem As Integer = 1
                                    For Each tbc As Object In tbd.Fields

                                        If IsNothing(Tabelas(tbd.name).ListaDeCampos) Then
                                            Tabelas(tbd.name).ListaDeCampos = New List(Of Campo)
                                        End If
                                        Tabelas(tbd.name).ListaDeCampos.Add(New Campo(tbd.Name, tbc.Name, CampoOrdem, TipoAccessToScript(tbc), "", "", "", "", ""))

                                        CampoOrdem += 1
                                    Next
                                End If

                                TabOrdem += 1
                            End If
                        Next
                    End If

                    ' trata comentários
                    If Importa And Criterios.Comentario Then

                        For Each tbd As Object In db.TableDefs
                            If Not (tbd.Name Like "MSys*" Or tbd.Name Like "~*" Or (TemNaLista(TabsSistema, tbd.name) And Not CType(Importa And Criterios.InfraSistema, Boolean)) Or (TemNaLista(TabsGerador, tbd.name) And Not CType(Importa And Criterios.InfraGerador, Boolean))) Then

                                ' antes dos campos, inclui linha de tabela
                                Dim Descr As DescrConcat = New DescrConcat(tbd.Name, Nothing, DaoPropTab(tbd, "description"))

                                ' comentários
                                Tabelas(tbd.name).Descr = Descr.Descr
                                Tabelas(tbd.name).Etiq = Descr.Etiq

                                For Each tbc As Object In tbd.Fields
                                    Dim DescrCampo As New DescrConcat(tbd.name, tbc.name, DaoPropCampo(tbc, "description"))
                                    Tabelas(tbd.name).Campos(tbc.name).Etiq = DescrCampo.Etiq
                                    Tabelas(tbd.name).Campos(tbc.name).Descr = DescrCampo.Descr
                                Next


                            End If
                        Next

                    End If

                    ' relacionamentos
                    If Importa And Criterios.Relacionamento Then
                        For Each tbr As Object In db.Relations

                            Dim Tabela_1 As String = tbr.Table
                            Dim Tabela_N As String = tbr.ForeignTable

                            If Not IsNothing(Tabelas(Tabela_1)) And Not IsNothing(Tabelas(Tabela_N)) Then

                                ' busca campos do relacionamento
                                Dim Campos_1 As String = ""
                                Dim Campos_N As String = ""

                                'Loop para carregar chaves primárias e estrangeiras
                                For Each fld As Object In tbr.Fields
                                    Campos_1 &= IIf(Campos_1 <> "", ";", "") & fld.Name
                                    Campos_N &= IIf(Campos_N <> "", ";", "") & fld.ForeignName
                                Next
                                Dim CascadeDel As String = IIf(CType(tbr.Attributes And DAO_RelationAttributeEnum_dbRelationDeleteCascade, Boolean), "CASCADE", "")
                                Dim CascadeUpd As String = IIf(CType(tbr.Attributes And DAO_RelationAttributeEnum_dbRelationUpdateCascade, Boolean), "CASCADE", "")
                                Dim Obrig As Boolean = Not CType(tbr.Attributes And DAO_RelationAttributeEnum_dbRelationDontEnforce, Boolean)
                                ListaDeRels.Add(New Gerador.Rel(NomeRel(Tabela_1, Tabela_N, ListaDeRels), Tabela_1, Campos_1, Tabela_N, Campos_N, CascadeDel, CascadeUpd, Obrig))
                            End If
                        Next
                    End If

                    ListaDeRels = ListaDeRels
                    db.Close()

                Catch ex As Exception
                    db.Close()
                    Throw ex
                End Try
            End Sub
            Public Sub CarregaOracle(ByVal Esquema As String, ByVal ConnStr As String)
                Dim Provider As String = Oracle

                GaranteExistenciaDasListas()

                ' classe
                Dim ListaDeClasses As New List(Of Gerador.Classe)
                If Importa And Criterios.Classe Then
                    ListaDeClasses.Add(New Gerador.Classe(Esquema, Esquema))
                End If
                Me.ListaDeClasses = ListaDeClasses

                'Carrega Tabelas
                Dim ListaDeTabelas As New List(Of Gerador.Tabela)
                Dim DS As DataSet = DSCarrega("SELECT TABLE_NAME FROM ALL_TABLES WHERE UPPER(OWNER) = :DB_NAME", ConnStr, ":DB_NAME", Esquema.ToUpper)
                If Importa And Criterios.Tabela Then

                    Dim TabOrdem As Integer = 1

                    For Each row As DataRow In DS.Tables(0).Rows
                        If Not row("TABLE_NAME") Like "BIN$*" And Not (Not (Importa And Criterios.InfraSistema) AndAlso Icraft.IcftBase.TemNaLista(TabsSistema, row("TABLE_NAME"))) Then

                            ' chave prima
                            Dim DSPrima As DataSet = DSCarrega("SELECT CC.COLUMN_NAME AS Chave_Prima, CC.TABLE_NAME, AC.CONSTRAINT_NAME FROM ALL_CONSTRAINTS AC, ALL_CONS_COLUMNS CC WHERE CC.CONSTRAINT_NAME = AC.CONSTRAINT_NAME AND CC.OWNER = AC.OWNER AND CC.TABLE_NAME = AC.TABLE_NAME AND AC.CONSTRAINT_TYPE = 'P' AND CC.TABLE_NAME = :TABELA AND UPPER(CC.OWNER) = :NOME ORDER BY CC.CONSTRAINT_NAME ASC", ConnStr, ":TABELA", row("TABLE_NAME"), ":NOME", Esquema)
                            Dim Chave_Prima As String = ""
                            Try
                                For Each fld As DataRow In DSPrima.Tables(0).Rows
                                    Chave_Prima &= IIf(Chave_Prima <> "", ";", "") & fld("Chave_Prima")
                                Next
                            Catch
                            End Try

                            ' campos
                            Dim ListaDeCampos As New List(Of Gerador.Campo)
                            If Importa And Criterios.Campo Then
                                Dim DSCampos As DataSet = DSCarrega("SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE FROM ALL_TAB_COLUMNS WHERE UPPER(OWNER) = :DB_NAME AND UPPER(TABLE_NAME)=:TABELA ORDER BY COLUMN_ID", ConnStr, ":DB_NAME", Esquema, ":TABELA", row("TABLE_NAME").ToString.ToUpper)
                                Dim CampoOrdem As Integer = 1

                                For Each tbc As DataRow In DSCampos.Tables(0).Rows
                                    ListaDeCampos.Add(New Campo(row("TABLE_NAME"), tbc("COLUMN_NAME"), CampoOrdem, "", TipoOracleToScript(tbc("DATA_TYPE"), tbc("DATA_LENGTH"), NZ(tbc("DATA_PRECISION"), ""), NZ(tbc("DATA_SCALE"), "")), "", tbc("COLUMN_NAME"), "", ""))
                                    CampoOrdem += 1
                                Next
                            End If

                            ListaDeTabelas.Add(New Tabela(row("TABLE_NAME"), TabOrdem, Chave_Prima, ListaDeCampos, row("TABLE_NAME"), Esquema, "", row("TABLE_NAME")))
                            TabOrdem += 1
                        End If
                    Next
                End If
                Me.ListaDeTabelas = ListaDeTabelas


                ' relacionamentos
                Dim ListaDeRels As New List(Of Gerador.Rel)
                If Importa And Criterios.Relacionamento Then
                    Dim DSRels As DataSet = DSCarrega("SELECT RR.CONSTRAINT_NAME CONSTRN, RR.TABLE_NAME TABN, R1.CONSTRAINT_NAME CONSTR1, R1.TABLE_NAME TAB1 FROM ALL_CONSTRAINTS RR, ALL_CONSTRAINTS R1 WHERE RR.CONSTRAINT_TYPE='R' AND RR.R_OWNER = R1.OWNER AND RR.R_CONSTRAINT_NAME = R1.CONSTRAINT_NAME AND RR.OWNER = :NOME ORDER BY TAB1 ASC", ConnStr, ":NOME", Esquema)

                    For Each tbr As DataRow In DSRels.Tables(0).Rows

                        Dim Tabela_1 As String = tbr("TAB1")
                        Dim Tabela_N As String = tbr("TABN")

                        ' busca campos do relacionamento
                        Dim Campos_1 As String = ""
                        Dim Campos_N As String = ""

                        'Loop para carregar chaves primárias e estrangeiras
                        'Monta SQL para carregar campos relacionados
                        Dim SQL As String = "SELECT * FROM (SELECT RR.OWNERN, RR.CONSTN, RN.POSITION POS, RN.COLUMN_NAME COLN, " & _
                                            "R1.COLUMN_NAME COL1 FROM (SELECT RR.OWNER OWNERN, RR.CONSTRAINT_NAME CONSTN, " & _
                                            "R1.OWNER OWNER1, R1.CONSTRAINT_NAME CONST1 FROM ALL_CONSTRAINTS RR, ALL_CONSTRAINTS R1 " & _
                                            "WHERE RR.CONSTRAINT_TYPE='R' AND RR.R_OWNER = R1.OWNER AND RR.R_CONSTRAINT_NAME = R1.CONSTRAINT_NAME " & _
                                            "AND RR.OWNER = :NOME) RR,(SELECT OWNER, CONSTRAINT_NAME, COLUMN_NAME, POSITION FROM ALL_CONS_COLUMNS) RN, " & _
                                            "(SELECT OWNER, CONSTRAINT_NAME, COLUMN_NAME, POSITION FROM ALL_CONS_COLUMNS) R1 WHERE RR.OWNERN = RN.OWNER " & _
                                            "AND RR.CONSTN = RN.CONSTRAINT_NAME AND RR.OWNER1 = R1.OWNER AND RR.CONST1 = R1.CONSTRAINT_NAME AND " & _
                                            "RN.POSITION = R1.POSITION ORDER BY RN.OWNER, RN.CONSTRAINT_NAME, RN.POSITION) WHERE " & _
                                            "OWNERN = :NOME AND CONSTN = :CONS"
                        Dim DSCamposRel As DataSet = DSCarrega(SQL, ConnStr, ":NOME", Esquema, ":CONS", tbr("CONSTRN"))
                        For Each row As DataRow In DSCamposRel.Tables(0).Rows
                            Campos_1 &= IIf(Campos_1 <> "", ";", "") & row("COL1")
                            Campos_N &= IIf(Campos_N <> "", ";", "") & row("COLN")
                        Next
                        Dim CascadeDel As String = "CASCADE"
                        Dim CascadeUpd As String = "CASCADE"
                        Dim Obrig As Boolean = True
                        ListaDeRels.Add(New Gerador.Rel(NomeRel(Tabela_1, Tabela_N, ListaDeRels), Tabela_1, Campos_1, Tabela_N, Campos_N, CascadeDel, CascadeUpd, Obrig))
                    Next
                End If
                Me.ListaDeRels = ListaDeRels
            End Sub
            Public Sub CarregaOracle(ByVal Esquema As String, ByVal Servico As String, ByVal Usuario As String, ByVal Senha As String)
                'Abre Conexão
                Dim ConnStr As String = "Data Source=[:VALOR.SERVICO];Persist Security Info=True;User ID=[:VALOR.USUARIO];Password=[:VALOR.SENHA];Unicode=True"
                CarregaOracle(Esquema, ConnStr)
                MacroSubstSQL(ConnStr, Nothing, ":SERVICO", Servico, ":USUARIO", Usuario, ":SENHA", Senha)
            End Sub
            Public Sub CarregaMySQL(ByVal Maquina As String, ByVal BancoDeDados As String, ByVal Usuario As String, ByVal Senha As String)

                ' abre conexão
                Dim ConnStr As String = "Server=[:VALOR.MAQUINA];Database=[:VALOR.BANCODADOS];Uid=[:VALOR.USUARIO];Pwd=[:VALOR.SENHA];"
                Dim Provider As String = MySQL
                MacroSubstSQL(ConnStr, Nothing, ":MAQUINA", Maquina, ":BANCODADOS", BancoDeDados, ":USUARIO", Usuario, ":SENHA", Senha)
                Dim Conf As New System.Configuration.ConnectionStringSettings("DIVERS_TRANSF_ORIGEM", ConnStr, Provider)

                ' classe
                Dim ListaDeClasses As New List(Of Gerador.Classe)
                ListaDeClasses.Add(New Gerador.Classe(BancoDeDados, BancoDeDados))
                Me.ListaDeClasses = ListaDeClasses

                'Carrega tabelas
                Dim DS As DataSet = DSCarrega("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE UPPER(TABLE_SCHEMA) = :DB_NAME", Conf, ":DB_NAME", BancoDeDados)
                Dim ListaDeTabelas As New List(Of Gerador.Tabela)
                Dim TabOrdem As Integer = 1

                For Each row As DataRow In DS.Tables(0).Rows
                    If Not row("TABLE_NAME") Like "BIN$*" And Not (Not (Importa And Criterios.InfraSistema) AndAlso Icraft.IcftBase.TemNaLista(TabsSistema, row("TABLE_NAME"))) Then

                        ' chave prima
                        Dim DSPrima As DataSet = DSCarrega("SELECT COLUMN_NAME AS Chave_Prima FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE CONSTRAINT_NAME ='PRIMARY' AND TABLE_NAME=:TABELA AND TABLE_SCHEMA=:ESQUEMA", Conf, ":TABELA", row("TABLE_NAME"), ":ESQUEMA", BancoDeDados)
                        Dim Chave_Prima As String = ""
                        Try
                            For Each fld As DataRow In DSPrima.Tables(0).Rows
                                Chave_Prima &= IIf(Chave_Prima <> "", ";", "") & fld("Chave_Prima")
                            Next
                        Catch
                        End Try

                        ' campos
                        Dim DSCampos As DataSet = DSCarrega("SELECT COLUMN_NAME,COLUMN_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE UPPER(TABLE_SCHEMA) = :DB_NAME AND UPPER(TABLE_NAME)=:TABELA", Conf, ":DB_NAME", BancoDeDados, ":TABELA", row("TABLE_NAME").ToString.ToUpper)
                        Dim ListaDeCampos As New List(Of Gerador.Campo)
                        Dim CampoOrdem As Integer = 1

                        For Each tbc As DataRow In DSCampos.Tables(0).Rows

                            ListaDeCampos.Add(New Campo(row("TABLE_NAME"), tbc("COLUMN_NAME"), CampoOrdem, "", "", tbc("COLUMN_TYPE"), tbc("COLUMN_NAME"), "", ""))
                            CampoOrdem += 1
                        Next

                        ListaDeTabelas.Add(New Tabela(row("TABLE_NAME"), TabOrdem, Chave_Prima, ListaDeCampos, row("TABLE_NAME"), BancoDeDados, "", row("TABLE_NAME")))
                        TabOrdem += 1
                    End If
                Next
                Me.ListaDeTabelas = ListaDeTabelas

                ' relacionamentos

                Dim NomeR As String = ""
                Dim Tabela_1 As String = ""
                Dim Tabela_N As String = ""
                Dim Campos_1 As String = ""
                Dim Campos_N As String = ""

                Dim DSRels As DataSet = DSCarrega("SELECT table_name as TABN, constraint_name as NOMEREL, column_name as COLN, referenced_table_name as TAB1, referenced_column_name as COL1, ordinal_position as pos FROM information_schema.KEY_COLUMN_USAGE K where(Not referenced_table_name Is null and table_schema=:DB_NAME) order by table_name,ordinal_position", Conf, ":DB_NAME", BancoDeDados)

                Try
                    Dim ListaDeRels As New List(Of Gerador.Rel)
                    For Each tbr As DataRow In DSRels.Tables(0).Rows

                        If NomeR <> tbr("TABN") & tbr("NOMEREL") Then
                            ' tem tab para gravar
                            If Tabela_N <> "" Then
                                ListaDeRels.Add(New Gerador.Rel(NomeRel(Tabela_1, Tabela_N, ListaDeRels), Tabela_1, Campos_1, Tabela_N, Campos_N, "CASCADE", "CASCADE", True))
                            End If

                            ' define
                            NomeR = tbr("TABN") & tbr("NOMEREL")
                            Tabela_1 = tbr("TAB1")
                            Tabela_N = tbr("TABN")
                            Campos_1 = ""
                            Campos_N = ""
                        End If


                        ' busca campos do relacionamento
                        Campos_1 &= IIf(Campos_1 <> "", ";", "") & tbr("COL1")
                        Campos_N &= IIf(Campos_N <> "", ";", "") & tbr("COLN")
                    Next
                    ' tem tab para gravar
                    If Tabela_N <> "" Then
                        ListaDeRels.Add(New Gerador.Rel(NomeRel(Tabela_1, Tabela_N, ListaDeRels), Tabela_1, Campos_1, Tabela_N, Campos_N, "CASCADE", "CASCADE", True))
                    End If

                    Me.ListaDeRels = ListaDeRels

                Catch ex As Exception
                    Throw New Exception("Captura de relacionamento (nome = " & NomeR & " / tabela 1 = " & Tabela_1 & " / tabela n = " & Tabela_N & ") apresentou problemas: " & ex.Message & ". Verifique permissões!")
                End Try
            End Sub
            Public Sub GravaGerador(Optional ByVal SistemaNovo As String = Nothing, Optional ByVal StrGerador As String = "strgerador")
                ' sobreposição de nome
                If Not IsNothing(SistemaNovo) Then
                    Me.Sistema = SistemaNovo
                End If

                ' verifica se já existe
                If DSValor("count(*)", "SISTEMA", StrGerador, "NOME=:SISNOME", ":SISNOME", Me.Sistema) > 0 Then
                    Throw New Exception("Sistema informado como destino já existe.")
                    Exit Sub
                End If

                ' insere sistema no gerador
                DSGrava("insert into sistema (nome,descr,ver,data_ver, rev) values(:nome,:descr,:ver,:data_ver,:rev)", "strgerador", ":nome", Me.Sistema, ":descr", NZ(Me.Descr, ""), ":ver", NZ(Me.Ver, ""), ":data_ver", NZ(Me.Data_Ver, Convert.DBNull), ":rev", NZ(Me.Rev, ""))

                ' insere classe no gerador
                If IsNothing(Me.ListaDeClasses) OrElse Me.ListaDeClasses.Count = 0 Then
                    DSGrava("insert into classe_OBJ (sistema, classe) values(:sistema,:classe)", StrGerador, ":sistema", Me.Sistema, ":classe", Me.Sistema)
                Else
                    For Each Cl As Classe In Me.ListaDeClasses
                        DSGrava("insert into classe_OBJ (sistema, classe, descr) values(:sistema,:classe, :descr)", StrGerador, ":sistema", Me.Sistema, ":classe", Cl("Classe"), ":descr", Cl("descr"))
                    Next
                End If

                ' insere tabela
                If Not IsNothing(Me.ListaDeTabelas) AndAlso Me.ListaDeTabelas.Count > 0 Then
                    For Each tab As Tabela In Me.ListaDeTabelas
                        DSGrava("insert into tabela (sistema, tabela, ordem, codigo, etiq, descr,classe,Chave_Prima) values(:sistema,:tabela,:ordem,:codigo, :etiq, :descr,:classe,:Chave_Prima)", "strgerador", ":sistema", Me.Sistema, ":tabela", tab("Tabela"), ":codigo", tab("Codigo"), ":ordem", tab("Ordem"), ":etiq", tab("etiq"), ":descr", tab("descr"), ":classe", tab("classe"), ":Chave_Prima", tab("Chave_Prima"))

                        ' insere campos
                        For Each camp As Campo In tab("Campos")
                            DSGrava("insert into campo(sistema, tabela, campo, ordem, etiq, descr, tipo_oracle, tipo_access, tipo_mysql, prop_extend) values(:sistema, :tabela, :campo, :ordem, :etiq, :descr, :tipo_oracle, :tipo_access, :tipo_mysql, :prop_extend)", "strgerador", ":sistema", Me.Sistema, ":tabela", tab("Tabela"), ":campo", camp("Campo"), ":ordem", camp("Ordem"), ":etiq", camp("Etiq"), ":descr", camp("Descr"), ":tipo_oracle", camp("Tipo_Oracle"), ":tipo_access", camp("Tipo_Access"), ":tipo_mysql", camp("Tipo_MySQL"), ":prop_extend", camp("prop_extend"))
                        Next
                    Next
                End If

                ' insere relacionamentos
                If Not IsNothing(Me.ListaDeRels) AndAlso Me.ListaDeRels.Count > 0 Then
                    For Each rr As Rel In Me.ListaDeRels
                        DSGrava("insert into relacionamento (sistema, nome, tabela_1, campo_1, tabela_n, campo_n, delete_cascade, update_cascade, obrig) values(:sistema, :nome, :tabela_1, :campo_1, :tabela_n, :campo_n, :delete_cascade, :update_cascade, :obrig)", "strgerador", ":sistema", Me.Sistema, ":nome", rr("Nome"), ":tabela_1", rr("Tabela_1"), ":campo_1", rr("Campo_1"), ":tabela_n", rr("Tabela_N"), ":campo_n", rr("Campo_N"), ":delete_cascade", NZV(rr("Delete_Cascade"), "CASCADE"), ":update_cascade", NZV(rr("Update_Cascade"), "CASCADE"), ":obrig", rr("Obrig"))
                    Next
                End If
            End Sub
            Property DaoPropTab(ByVal TbDef As Object, ByVal NomeProp As String) As Object
                Get
                    Try
                        Return TbDef.Properties(NomeProp).Value
                    Catch
                    End Try
                    Return ""
                End Get
                Set(ByVal value As Object)
                    Try
                        If NZ(value, "") = "" Then
                            TbDef.Properties.Delete(NomeProp)
                        Else
                            Try
                                TbDef.Properties(NomeProp).Value = value
                            Catch
                                Dim Prop As Object = TbDef.CreateProperty(NomeProp, DAO_DataTypeEnum_dbText, value)
                                TbDef.Properties.Append(Prop)
                            End Try
                        End If
                    Catch
                    End Try
                End Set
            End Property
            Property DaoPropCampo(ByVal Fld As Object, ByVal NomeProp As String) As Object
                Get
                    Try
                        Return Fld.Properties(NomeProp).Value
                    Catch
                    End Try
                    Return ""
                End Get
                Set(ByVal value As Object)
                    Try
                        If NZ(value, "") = "" Then
                            Fld.Properties.Delete(NomeProp)
                        Else
                            Try
                                Fld.Properties(NomeProp).Value = value
                            Catch
                                Dim Prop As Object = Fld.CreateProperty(NomeProp, DAO_DataTypeEnum_dbText, value)
                                Fld.Properties.Append(Prop)
                            End Try
                        End If
                    Catch
                    End Try
                End Set
            End Property

            Function Copia() As Gerador
                Dim NGera As New Gerador

                NGera.ListaDeTabelas = New List(Of Tabela)
                If Not IsNothing(ListaDeTabelas) Then
                    NGera.ListaDeTabelas.AddRange(ListaDeTabelas)
                End If

                NGera.ListaDeRels = New List(Of Rel)
                If Not IsNothing(ListaDeRels) Then
                    NGera.ListaDeRels.AddRange(ListaDeRels)
                End If

                NGera.ListaDeClasses = New List(Of Classe)
                If Not IsNothing(ListaDeClasses) Then
                    NGera.ListaDeClasses.AddRange(ListaDeClasses)
                End If

                NGera.ListaDeVisoes = New List(Of Visao)
                If Not IsNothing(ListaDeVisoes) Then
                    NGera.ListaDeVisoes.AddRange(ListaDeVisoes)
                End If

                NGera.ListaDeObjs = New List(Of Obj)
                If Not IsNothing(ListaDeObjs) Then
                    NGera.ListaDeObjs.AddRange(ListaDeObjs)
                End If

                NGera.ListaDeUsuarios = New List(Of Usuario)
                If Not IsNothing(ListaDeUsuarios) Then
                    NGera.ListaDeUsuarios.AddRange(ListaDeUsuarios)
                End If

                NGera.ListaDeDireitos = New List(Of Direito)
                If Not IsNothing(ListaDeDireitos) Then
                    NGera.ListaDeDireitos.AddRange(ListaDeDireitos)
                End If

                NGera.Data_Ver = Data_Ver
                NGera.Ver = Ver
                NGera.Tipo = Tipo
                NGera.Sistema = Sistema
                NGera.Descr = Descr
                NGera.Importa = Importa
                NGera.Exporta = Exporta


                Return NGera
            End Function

            Public Sub GravaMSAccess(ByVal Pagina As Page, ByVal ArquivoMDB As String)
                Dim NGera As Gerador = Me.Copia

                Dim db As Object = Nothing
                ArquivoMDB = FileExpr(ArquivoMDB)

                Dim de As Object
                Try
                    de = CreateObject("DAO.DBEngine.36")
                Catch ex As Exception
                    Throw New Exception("Necessária inclusão de referência para DBEngine.36.")
                End Try

                If System.IO.File.Exists(ArquivoMDB) Then
                    If NGera.Exporta And Criterios.Iniciar Then
                        Kill(ArquivoMDB)
                    Else
                        db = de(0).OpenDatabase(ArquivoMDB)
                    End If
                End If
                If Not System.IO.File.Exists(ArquivoMDB) Then
                    db = de(0).CreateDatabase(ArquivoMDB, DAO_LanguageConstants_dbLangGeneral)
                End If

                Try

                    ' TABELAS DO SISTEMA
                    If NGera.Exporta And Criterios.InfraSistema Then
                        NGera.Importa = NGera.Importa - (NGera.Importa And Criterios.Iniciar)
                        NGera.CarregaXML("~/uc/icftgera/sistema.xml")
                    End If

                    ' Tabelas do gerador
                    If NGera.Exporta And Criterios.InfraGerador Then
                        NGera.Importa = NGera.Importa - (NGera.Importa And Criterios.Iniciar)
                        NGera.CarregaXML("~/uc/icftgera/gerador.xml")
                    End If

                    ' insere tabela
                    If NGera.Exporta And Criterios.Tabela Then
                        For Each tab As Tabela In NGera.ListaDeTabelas

                            Dim tb As Object = Nothing
                            tb = de(0).databases(0).CreateTableDef(tab("Tabela"))

                            If NGera.Exporta And Criterios.Campo Then
                                ' insere campos
                                For Each camp As Campo In tab("Campos")
                                    Dim m As Match = RegexMatches(TipoScriptToAccess(camp), "(.*);([^ \(]*)( \((.*)\))*")

                                    Dim cp As Object = tb.CreateField(camp("Campo"), IIf(m.Groups(2).Value = "20", "4", m.Groups(2).Value))
                                    Try
                                        cp.AllowZeroLength = True
                                    Catch
                                    End Try

                                    If m.Groups(4).Value <> "" Then
                                        cp.Size = m.Groups(4).Value
                                    End If
                                    tb.Fields.Append(cp)
                                Next
                            End If

                            ' insere chave primária
                            If tab("chave_prima") <> "" Then
                                Dim chp As Object = tb.CreateIndex("PrimaryKey")
                                chp.Primary = True

                                For Each Campo As String In Split(tab("chave_prima"), ";")
                                    Dim fld As Object = chp.CreateField(Campo)
                                    chp.Fields.Append(fld)
                                Next
                                tb.Indexes.Append(chp)
                            End If

                            db.TableDefs.Append(tb)

                            If NGera.Exporta And Criterios.Relacionamento Then
                                ' define propriedades
                                DaoPropTab(tb, "Description") = New DescrConcat(tb.name, Nothing, tab("Classe"), tab("Etiq"), tab("Descr")).Texto
                                For Each camp As Campo In tab("Campos")

                                    DaoPropCampo(tb.Fields(camp("Campo")), "Description") = New DescrConcat(tab.Tabela, camp.Campo, Nothing, camp("Etiq"), camp("Descr")).Texto
                                Next
                            End If
                        Next
                    End If

                    If NGera.Exporta And Criterios.Relacionamento Then
                        ' insere relacionamentos
                        If Not IsNothing(NGera.ListaDeRels) Then
                            For Each rr As Rel In NGera.ListaDeRels
                                ' trata atributo de relacionamento
                                Dim p As Integer = 0
                                If rr("Delete_Cascade") = "CASCADE" Then
                                    p += DAO_RelationAttributeEnum_dbRelationDeleteCascade
                                End If
                                If rr("Update_Cascade") = "CASCADE" Then
                                    p += DAO_RelationAttributeEnum_dbRelationUpdateCascade
                                End If
                                If Not rr("OBRIG") Then
                                    p += DAO_RelationAttributeEnum_dbRelationDontEnforce
                                End If

                                Dim rl As Object = db.CreateRelation(rr("Nome"), rr("Tabela_1"), rr("Tabela_N"))
                                rl.Attributes = p

                                ' inclui campos
                                Dim cps As String() = Split(rr("Campo_1"), ";")
                                Dim cpsrel As String() = Split(rr("Campo_N"), ";")
                                For z As Integer = 0 To cps.Length - 1
                                    Dim cp As Object = rl.CreateField(cps(z))
                                    cp.ForeignName = cpsrel(z)
                                    rl.Fields.Append(cp)
                                Next


                                Try
                                    db.Relations.Append(rl)
                                Catch
                                    Dim rltcampo As String = ""
                                    Dim rltfcampo As String = ""
                                    For Each fld As Object In rl.Fields
                                        rltcampo &= IIf(rltcampo <> "", ";", "") & fld.Name
                                        rltfcampo &= IIf(rltfcampo <> "", ";", "") & fld.ForeignName
                                    Next
                                    ShowJSMessage(Pagina, "falta rel de " & rl.Table & " (" & rltcampo & ")" & " para " & rl.ForeignTable & "(" & rltfcampo & ").")
                                End Try
                            Next
                        End If
                    End If

                    If (NGera.Exporta And Criterios.InfraGerador) Then
                        If (NGera.Exporta And Criterios.DadosGerador) Then
                            db.execute(MacroSubstSQLText("insert into ger_sistema(nome,descr,ver,data_ver,rev,param,prop_extend) values([:nome],[:descr],[:ver],[:data_ver],[:rev],[:param],[:prop_extend])", ParamArrayToArrayList(":nome", NGera.Sistema, ":descr", NGera.Descr, ":ver", NGera.Ver, ":data_ver", NGera.Data_Ver, ":rev", NGera.Rev, ":param", "", ":prop_extend", "")), 128)
                            For Each CL As Classe In NGera.ListaDeClasses
                                db.execute(MacroSubstSQLText("insert into ger_classe(sistema,classe,descr) values([:sistema],[:classe],[:descr])", ParamArrayToArrayList(":sistema", NGera.Sistema, ":classe", CL.Classe, ":descr", CL.Descr)), 128)
                            Next
                            For Each TB As Tabela In NGera.ListaDeTabelas
                                db.execute(MacroSubstSQLText("insert into ger_tabela(sistema,tabela,ordem,codigo,etiq,descr,chave_prima,classe,chave_apres,depende,interfere,chave_filtro) values([:sistema],[:tabela],[:ordem],[:codigo],[:etiq],[:descr],[:chave_prima],[:classe],[:chave_apres],[:depende],[:interfere],[:chave_filtro])", ParamArrayToArrayList(":sistema", NGera.Sistema, ":tabela", TB.Tabela, ":ordem", TB.Ordem, ":codigo", TB.Codigo, ":etiq", TB.Etiq, ":descr", TB.Descr, ":chave_prima", TB.Chave_Prima, ":classe", TB.Classe, ":chave_apres", "", ":depende", 0, ":interfere", 0, ":chave_filtro", "")), 128)
                            Next
                            For Each CP As Campo In NGera.ListaDeCampos
                                db.execute(MacroSubstSQLText("insert into ger_campo(sistema,tabela,campo,ordem,etiq,descr,prop_extend,tipo_access,tipo_oracle,tipo_mysql,formato,valor_padrao,auto) values([:sistema],[:tabela],[:campo],[:ordem],[:etiq],[:descr],[:prop_extend],[:tipo_access],[:tipo_oracle],[:tipo_mysql],[:formato],[:valor_padrao],[:auto])", ParamArrayToArrayList(":sistema", NGera.Sistema, ":tabela", CP.Tabela, ":campo", CP.Campo, ":ordem", CP.Ordem, ":etiq", CP.Etiq, ":descr", CP.Descr, ":prop_extend", CP.Prop_Extend, ":tipo_access", CP.Tipo_Access, ":tipo_oracle", CP.Tipo_Oracle, ":tipo_mysql", CP.Tipo_MySQL, ":formato", "", ":valor_padrao", "", ":auto", "")), 128)
                            Next
                            For Each RL As Rel In NGera.ListaDeRels
                                db.EXECUTE(MacroSubstSQLText("insert into ger_relacionamento(sistema,nome,tabela_1,campo_1,chave_apres_1,tabela_n,campo_n,chave_apres_n,delete_cascade,update_cascade,obrig,chave_expr,formato,mascara,valor_padrao,compr_zero,auto,expr_apres) values([:sistema],[:nome],[:tabela_1],[:campo_1],[:chave_apres_1],[:tabela_n],[:campo_n],[:chave_apres_n],[:delete_cascade],[:update_cascade],[:obrig],[:chave_expr],[:formato],[:mascara],[:valor_padrao],[:compr_zero],[:auto],[:expr_apres])", ParamArrayToArrayList(":sistema", NGera.Sistema, ":nome", RL.Nome, ":tabela_1", RL.Tabela_1, ":campo_1", RL.Campo_1, ":chave_apres_1", "", ":tabela_n", RL.Tabela_N, ":campo_n", RL.Campo_N, ":chave_apres_n", "", ":delete_cascade", RL.Delete_Cascade, ":update_cascade", RL.Update_Cascade, ":obrig", RL.Obrig, ":chave_expr", "", ":formato", "", ":mascara", "", ":valor_padrao", "", ":compr_zero", True, ":auto", "", ":expr_apres", "")), 128)
                            Next
                        End If
                    End If
                    db.Close()

                Catch ex As Exception
                    db.close()
                    Throw ex
                End Try


            End Sub
            Public Sub GravaOracleSemRestr(ByVal S As System.IO.StreamWriter, ByVal ExecutaNoOracle As Boolean, ByVal Params As ArrayList)
                ' preparo do ambiente de criação
                S.WriteLine("/* **********************************************************************************")
                S.WriteLine("   INTERCRAFT SOLUTIONS INFORMÁTICA LTDA - 2008")
                S.WriteLine("   SCRIPT CRIADO A PARTIR DE http://www.intercraft.inf.br/gerador/basedados/transf_estrut.aspx")
                S.WriteLine("   EM " & Format(Now, "dd/MM/yyyy HH:mm"))
                S.WriteLine("")
                S.WriteLine("   IMPORTANTE!!!")
                S.WriteLine("   * O SISTEMA LOGARÁ SYSTEM PARA INICIAR A EXECUÇÃO, LOGO, MUITO CUIDADO!")
                S.WriteLine("   * CASO O USUÁRIO DO ESQUEMA JÁ EXISTA, ESTE SERÁ ELIMINADO, PORTANTO, TENHA A CERTEZA DAQUILO QUE ESTÁ FAZENDO.")
                S.WriteLine(MacroSubstSQLText("   * ARQUIVO C:\[:VALOR.ESQUEMA].LOG REGISTRARÁ O LOG DE EXECUÇÃO, QUE DEVE SER VERIFICADO.", Params))
                S.WriteLine("   * CASO BANCO POSSUA ACENTOS, PRECISARÁ GARANTIR FORMATO ANSI ANTES DA EXECUÇÃO. PARA ISSO, EDITE O CONTEÚDO NO NOTEPAD E SALVE FAZENDO OPÇÃO POR ESTE FORMATO.")
                S.WriteLine("*/")
                S.WriteLine("")
                S.WriteLine("SET ECHO ON")
                S.WriteLine(MacroSubstSQLText("SPOOL C:\[:VALOR.ESQUEMA]_PARTE1.LOG", Params))
                S.WriteLine("")
                S.WriteLine("-- GARANTINDO QUE TODOS TENHA ACESSO AOS RECURSOS DE PACOTE")
                S.WriteLine(MacroSubstSQLText("CONNECT [:VALOR.USUARIOSYS]/[:VALOR.SENHASYS]@[:VALOR.SERVICO] AS SYSDBA", Params))
                S.WriteLine("GRANT EXECUTE ON SYS.UTL_FILE TO PUBLIC;")
                S.WriteLine("GRANT EXECUTE ON SYS.UTL_SMTP TO PUBLIC;")
                S.WriteLine("GRANT EXECUTE ON SYS.UTL_TCP TO PUBLIC;")
                S.WriteLine("GRANT SELECT ON SYS.V_$SESSION TO PUBLIC;")
                S.WriteLine("GRANT QUERY REWRITE TO PUBLIC;")
                S.WriteLine("GRANT CREATE MATERIALIZED VIEW TO PUBLIC;")
                S.WriteLine("ALTER SESSION SET QUERY_REWRITE_ENABLED = TRUE;")

                S.WriteLine("")
                S.WriteLine("-- ELIMINA OS JOBS")
                S.WriteLine("begin")
                S.WriteLine(MacroSubstSQLText("for cur in (select job from user_jobs where schema_user='[:VALOR.ESQUEMA]') loop", Params))
                S.WriteLine("dbms_job.remove(cur.job);")
                S.WriteLine("end loop;")
                S.WriteLine("end;")
                S.WriteLine("/")

                S.WriteLine("")
                S.WriteLine("-- INICIALIZANDO TABLESPACE")
                S.WriteLine(MacroSubstSQLText("ALTER TABLESPACE T_[:VALOR.ESQUEMA]_DAT OFFLINE;", Params))
                S.WriteLine(MacroSubstSQLText("DROP TABLESPACE T_[:VALOR.ESQUEMA]_DAT INCLUDING CONTENTS;", Params))
                S.WriteLine(MacroSubstSQLText("CREATE TABLESPACE T_[:VALOR.ESQUEMA]_DAT DATAFILE '[:VALOR.DIRTABLESPACE]\[:VALOR.ESQUEMA].DBF' SIZE 100M REUSE", Params))
                S.WriteLine("AUTOEXTEND ON NEXT 50M MAXSIZE UNLIMITED EXTENT MANAGEMENT LOCAL;")
                S.WriteLine("")
                S.WriteLine("-- DEFININDO USUÁRIO DE ESQUEMA")
                S.WriteLine(MacroSubstSQLText("DROP USER [:VALOR.ESQUEMA] CASCADE;", Params))
                S.WriteLine(MacroSubstSQLText("CREATE USER [:VALOR.ESQUEMA] IDENTIFIED BY [:VALOR.SENHAESQUEMA] DEFAULT TABLESPACE T_[:VALOR.ESQUEMA]_DAT;", Params))
                S.WriteLine(MacroSubstSQLText("GRANT DBA TO [:VALOR.ESQUEMA];", Params))
                S.WriteLine(MacroSubstSQLText("GRANT ALL PRIVILEGES TO [:VALOR.ESQUEMA];", Params))
                S.WriteLine("")
                S.WriteLine("-- CONECTANDO COM USUÁRIO DE ESQUEMA")
                S.WriteLine("DISCONNECT;")
                S.WriteLine(MacroSubstSQLText("CONNECT [:VALOR.ESQUEMA]/[:VALOR.SENHAESQUEMA]@[:VALOR.SERVICO];", Params))
                S.WriteLine("")

                ' prepara tabelas
                S.WriteLine("/* **********************************************************************************")
                S.WriteLine("   CRIAÇÃO DE TABELAS")
                S.WriteLine("*/")
                S.WriteLine("")

                For Each tab As Tabela In Me.ListaDeTabelas
                    If NZ(tab("Codigo"), "") <> "" And Not (Not (Importa And Criterios.InfraSistema) AndAlso Icraft.IcftBase.TemNaLista(TabsSistema, tab("codigo"))) Then
                        Dim ChavePrima As ArrayList = ParamArrayToArrayList(Split(tab("Chave_Prima"), ";"))
                        S.WriteLine("-- TABELA " & tab("Tabela"))
                        S.WriteLine(MacroSubstSQLText("CREATE TABLE [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & " (", Params))
                        Dim Texto As String = ""

                        For Each camp As Campo In tab("Campos")
                            Dim NotNull As Boolean = False
                            If ChavePrima.Contains(camp("Campo")) Then
                                NotNull = True
                            End If
                            Texto &= IIf(Texto <> "", "," & Chr(13) & Chr(10) & "   ", "") & SqlExpr(camp("Campo"), Chr(34)) & " "
                            Texto &= TipoScriptToOracle(camp)
                            'Texto &= IIf(NZ(camp("Valor_Padrao"), "") <> "", " DEFAULT '" & camp("Valor_Padrao") & "'", "")
                            Texto &= IIf(NotNull, " NOT NULL", "")
                        Next

                        S.WriteLine("   " & Texto & ",")
                        S.WriteLine("   SYS_MOMENTO_CRIA DATE,")
                        S.WriteLine("   SYS_USUARIO_CRIA VARCHAR2 (100),")
                        S.WriteLine("   SYS_LOCAL_CRIA VARCHAR2 (100),")
                        S.WriteLine("   SYS_MOMENTO_ATUALIZA DATE,")
                        S.WriteLine("   SYS_USUARIO_ATUALIZA VARCHAR2 (100),")
                        S.WriteLine("   SYS_LOCAL_ATUALIZA VARCHAR2 (100),")
                        S.WriteLine("   SYS_STATUS CHAR (1)")
                        S.WriteLine(");")
                        S.WriteLine("")

                        If ChavePrima.Count <> 0 Then
                            Dim Txt As String = ""

                            For Each Item As String In ChavePrima.ToArray
                                Txt &= IIf(Txt <> "", ", ", "") & SqlExpr(Item, Chr(34))
                            Next
                            S.WriteLine(MacroSubstSQLText("ALTER TABLE [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & " ADD CONSTRAINT ""ID_" & tab("Tabela") & """" & Chr(13) & Chr(10) & "PRIMARY KEY(" & Txt & ");", Params))
                            S.WriteLine("")
                        End If

                        S.WriteLine(MacroSubstSQLText("COMMENT ON TABLE [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & " IS '" & tab("Classe") & " | " & tab("Descr") & "';", Params))
                        S.WriteLine("")

                        For Each camp As Campo In tab("Campos")
                            S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & "." & SqlExpr(camp("Campo"), Chr(34)) & " IS '" & camp("Etiq") & " | " & camp("Descr") & "';", Params))
                        Next

                        S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & ".SYS_MOMENTO_CRIA IS 'Sys_Momento_Cria | Registra momento de gravação';", Params))
                        S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & ".SYS_USUARIO_CRIA IS 'Sys_Usuario_Cria | Registra usuário que gravou';", Params))
                        S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & ".SYS_LOCAL_CRIA IS 'Sys_Local_Cria | Registra local de gravação';", Params))
                        S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & ".SYS_MOMENTO_ATUALIZA IS 'Sys_Momento_Atualiza | Registra momento de atualização';", Params))
                        S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & ".SYS_USUARIO_ATUALIZA IS 'Sys_Usuario_Atualiza | Registra usuário que atualizou';", Params))
                        S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & ".SYS_LOCAL_ATUALIZA IS 'Sys_Local_Atualiza | Registra local de atualização';", Params))
                        S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA]." & SqlExpr(tab("Tabela"), Chr(34)) & ".SYS_STATUS IS 'Sys_Status | Registra o status';", Params))
                        S.WriteLine("")
                    End If

                    ' incluir rotina de indices
                Next

                If Exporta And Criterios.InfraSistema Then

                    ' cria tabelas de sistema
                    S.WriteLine("/* **********************************************************************************")
                    S.WriteLine("   CRIAÇÃO DE TABELAS DO SISTEMA")
                    S.WriteLine("*/")
                    S.WriteLine("")

                    S.WriteLine("-- TABELA DE SISTEMA - CONFIGURAÇÕES GLOBAIS")
                    S.WriteLine(MacroSubstSQLText("CREATE TABLE [:VALOR.ESQUEMA].SYS_CONFIG_GLOBAL (", Params))
                    S.WriteLine("   PARAM VARCHAR2 (70) NOT NULL,")
                    S.WriteLine("   DESCR VARCHAR2 (4000),")
                    S.WriteLine("   CONFIG VARCHAR2 (4000)")
                    S.WriteLine(");")
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("ALTER TABLE [:VALOR.ESQUEMA].SYS_CONFIG_GLOBAL ADD CONSTRAINT ID_CONFIG_GLOBAL PRIMARY KEY (PARAM);", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON TABLE [:VALOR.ESQUEMA].SYS_CONFIG_GLOBAL IS 'Sistema | Armazena parâmetros globais do aplicativo.';", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_CONFIG_GLOBAL.PARAM IS 'PARAM | Parametro';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_CONFIG_GLOBAL.DESCR IS 'DESCR | Descrição';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_CONFIG_GLOBAL.CONFIG IS 'CONFIG | Configuração';", Params))
                    S.WriteLine("")

                    S.WriteLine("-- TABELA DE SISTEMA - CONFIGURAÇÕES DE USUARIOS")
                    S.WriteLine(MacroSubstSQLText("CREATE TABLE [:VALOR.ESQUEMA].SYS_CONFIG_USUARIO (", Params))
                    S.WriteLine("   USUARIO VARCHAR2 (100),")
                    S.WriteLine("   PARAM VARCHAR2 (70),")
                    S.WriteLine("   DESCR VARCHAR2 (4000),")
                    S.WriteLine("   CONFIG VARCHAR2 (4000)")
                    S.WriteLine(");")
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("ALTER TABLE [:VALOR.ESQUEMA].SYS_CONFIG_USUARIO ADD CONSTRAINT ID_CONFIG_USUARIO PRIMARY KEY (USUARIO,PARAM);", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON TABLE [:VALOR.ESQUEMA].SYS_CONFIG_USUARIO IS 'Sistema | Registro de ocorrências do sistema.';", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_CONFIG_USUARIO.USUARIO IS 'USUARIO | Usuário';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_CONFIG_USUARIO.PARAM IS 'PARAM | Parametro';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_CONFIG_USUARIO.DESCR IS 'DESCR | Descrição';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_CONFIG_USUARIO.CONFIG IS 'CONFIG | Configuração';", Params))
                    S.WriteLine("")


                    S.WriteLine("-- TABELA DE SISTEMA - REGISTRO DE EXCLUSÕES")
                    S.WriteLine(MacroSubstSQLText("CREATE TABLE [:VALOR.ESQUEMA].SYS_DELETE (", Params))
                    S.WriteLine("   TABELA VARCHAR2 (70),")
                    S.WriteLine("   CHAVE VARCHAR2 (300),")
                    S.WriteLine("   MOMENTO DATE,")
                    S.WriteLine("   USUARIO VARCHAR2 (100),")
                    S.WriteLine("   LOCAL VARCHAR2 (20)")
                    S.WriteLine(");")
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("ALTER TABLE [:VALOR.ESQUEMA].SYS_DELETE ADD CONSTRAINT ID_DELETE PRIMARY KEY (TABELA, CHAVE);", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON TABLE [:VALOR.ESQUEMA].SYS_DELETE IS 'Sistema | Registro de exclusões do sistema.';", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_DELETE.TABELA IS 'TABELA | Tabela';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_DELETE.CHAVE IS 'CHAVE | Chave';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_DELETE.MOMENTO IS 'MOMENTO | Momento';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_DELETE.USUARIO IS 'USUARIO | Usuário';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_DELETE.LOCAL IS 'LOCAL | Local';", Params))
                    S.WriteLine("")

                    S.WriteLine("-- TABELA DE SISTEMA - CONFIGURAÇÕES DE LOCALIDADE")
                    S.WriteLine(MacroSubstSQLText("CREATE TABLE [:VALOR.ESQUEMA].SYS_LOCALID (", Params))
                    S.WriteLine("   NOME VARCHAR2 (20) NOT NULL,")
                    S.WriteLine("   CORRENTE INTEGER,")
                    S.WriteLine("   PACOTE INTEGER,")
                    S.WriteLine("   PACOTE_REC INTEGER,")
                    S.WriteLine("   MOMENTO DATE,")
                    S.WriteLine("   MOMENTO_REC DATE,")
                    S.WriteLine("   MODELO INTEGER,")
                    S.WriteLine("   OBS VARCHAR2 (300)")
                    S.WriteLine(");")
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("ALTER TABLE [:VALOR.ESQUEMA].SYS_LOCALID ADD CONSTRAINT ID_SYS_LOCAL PRIMARY KEY (NOME);", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON TABLE [:VALOR.ESQUEMA].SYS_LOCALID IS 'Sistema | Especificações de localidade para sistema distribuído.';", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_LOCALID.NOME IS 'NOME | Nome';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_LOCALID.CORRENTE IS 'CORRENTE | Corrente';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_LOCALID.PACOTE IS 'PACOTE | Pacote';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_LOCALID.PACOTE_REC IS 'PACOTE_REC | Pacote Recebido';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_LOCALID.MOMENTO IS 'MOMENTO | Momento';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_LOCALID.MOMENTO_REC IS 'MOMENTO_REC | Momento Recebido';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_LOCALID.MODELO IS 'MODELO | Modelo';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_LOCALID.OBS IS 'OBS | Observação';", Params))
                    S.WriteLine("")

                    S.WriteLine("-- TABELA DE SISTEMA - REGISTRO DE OCORRÊNCIAS")
                    S.WriteLine(MacroSubstSQLText("CREATE TABLE [:VALOR.ESQUEMA].SYS_OCORRENCIA (", Params))
                    S.WriteLine("   SEQ INTEGER NOT NULL,")
                    S.WriteLine("   APLICACAO VARCHAR2 (3000),")
                    S.WriteLine("   OCORRENCIA VARCHAR2 (3000),")
                    S.WriteLine("   USUARIO VARCHAR2 (100),")
                    S.WriteLine("   MOMENTO DATE,")
                    S.WriteLine("   LOCAL VARCHAR2 (20)")
                    S.WriteLine(");")
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("ALTER TABLE [:VALOR.ESQUEMA].SYS_OCORRENCIA ADD CONSTRAINT ID_OCORRENCIA PRIMARY KEY (SEQ);", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON TABLE [:VALOR.ESQUEMA].SYS_OCORRENCIA IS 'Sistema | Registro de ocorrências do sistema.';", Params))
                    S.WriteLine("")
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_OCORRENCIA.SEQ IS 'SEQ | Sequencial';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_OCORRENCIA.APLICACAO IS 'APLICACAO | Aplicação';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_OCORRENCIA.OCORRENCIA IS 'OCORRENCIA | Ocorrência';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_OCORRENCIA.USUARIO IS 'USUARIO | Usuário';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_OCORRENCIA.MOMENTO IS 'MOMENTO | Momento';", Params))
                    S.WriteLine(MacroSubstSQLText("COMMENT ON COLUMN [:VALOR.ESQUEMA].SYS_OCORRENCIA.LOCAL IS 'LOCAL | Local';", Params))
                    S.WriteLine("")
                End If

                S.WriteLine("SPOOL OFF")
            End Sub
            Public Sub GravaOracleRestr(ByVal S As System.IO.StreamWriter, ByVal ExecutaNoOracle As Boolean, ByVal Params As ArrayList)
                ' prepara constraints
                S.WriteLine("/* **********************************************************************************")
                S.WriteLine("   CRIAÇÃO DE RELACIONAMENTOS")
                S.WriteLine("*/")
                S.WriteLine("")
                S.WriteLine("SET ECHO ON")
                S.WriteLine(MacroSubstSQLText("SPOOL C:\[:VALOR.ESQUEMA]_PARTE2.LOG", Params))
                S.WriteLine(MacroSubstSQLText("CONNECT [:VALOR.ESQUEMA]/[:VALOR.SENHAESQUEMA]@[:VALOR.SERVICO];", Params))
                S.WriteLine("")
                For Each rr As Rel In Me.ListaDeRels
                    Dim CampoN As String = ""
                    For Each Item As String In Split(rr("Campo_N"), ";")
                        CampoN &= IIf(CampoN <> "", ", ", "") & SqlExpr(Item, Chr(34))
                    Next
                    Dim Campo1 As String = ""
                    For Each Item As String In Split(rr("Campo_1"), ";")
                        Campo1 &= IIf(Campo1 <> "", ", ", "") & SqlExpr(Item, Chr(34))
                    Next
                    S.WriteLine(MacroSubstSQLText("ALTER TABLE [:VALOR.ESQUEMA]." & SqlExpr(rr("Tabela_N"), Chr(34)) & " ADD CONSTRAINT " & SqlExpr(rr("Nome"), Chr(34)) & " FOREIGN KEY(" & CampoN & ") REFERENCES [:VALOR.ESQUEMA]." & SqlExpr(rr("Tabela_1"), Chr(34)) & "(" & Campo1 & ");", Params))
                Next
                S.WriteLine("")

                ' cria os triggers
                S.WriteLine("/* **********************************************************************************")
                S.WriteLine("   CRIAÇÃO DE TRIGGERS PARA REGISTRO DE INCLUSÃO, ATUALIZAÇÃO E EXCLUSÃO")
                S.WriteLine("*/")
                S.WriteLine("")
                For Each tab As Tabela In Me.ListaDeTabelas
                    Dim ChavePrima As String = Mid(Replace(";" & tab("Chave_Prima"), ";", " || :OLD."), 5)
                    If NZ(tab("Codigo"), "") <> "" Then
                        S.WriteLine(MacroSubstSQLText("CREATE OR REPLACE TRIGGER [:VALOR.ESQUEMA].BEF_" & tab("Tabela") & " BEFORE UPDATE OR INSERT OR DELETE ON [:VALOR.ESQUEMA]." & tab("TAbela") & " FOR EACH ROW", Params))
                        S.WriteLine("DECLARE")
                        S.WriteLine("   TOT_DEL INTEGER;")
                        S.WriteLine("   -- variáveis que preencherão o retorno das consultas")
                        S.WriteLine("   CONN_USER VARCHAR2(100);")
                        S.WriteLine("   CONN_IP VARCHAR2(100);")
                        S.WriteLine("   CONN_MACHINE VARCHAR2(100);")
                        S.WriteLine("BEGIN")
                        S.WriteLine("   IF LPAD(USER,5) <> 'REPL_' THEN")

                        S.WriteLine("      -- Consulta que retornará o HOST, IP, USUÁRIO que acessou via internet")
                        S.WriteLine("      select module, client_info, action into conn_machine, conn_ip, conn_user from v$session where audsid = userenv('sessionid');")

                        S.WriteLine("      -- No caso da variável CONN_IP ser nula, significará que o acesso está sendo feito localmente")
                        S.WriteLine("      If (conn_ip Is null) Then")
                        S.WriteLine("         -- Consulta que retornará o HOST, IP, USUÁRIO que acessou via CIAD")
                        S.WriteLine("         select sys_context('userenv','host'), sys_context('userenv','ip_address'), sys_context('userenv','session_user') into conn_user, conn_ip, conn_machine from dual;")
                        S.WriteLine("      end if;")

                        S.WriteLine("      IF DELETING Then")
                        S.WriteLine(MacroSubstSQLText("         SELECT COUNT(*) INTO TOT_DEL FROM [:VALOR.ESQUEMA].SYS_DELETE WHERE TABELA = '" & tab("Tabela") & "' AND CHAVE || '' = " & ChavePrima & " || '';", Params))
                        S.WriteLine("         IF TOT_DEL = 0 THEN")
                        S.WriteLine(MacroSubstSQLText("            INSERT INTO [:VALOR.ESQUEMA].SYS_DELETE VALUES ('" & tab("Tabela") & "', " & ChavePrima & ", SYSDATE, USER, '[:VALOR.ESQUEMA]');", Params))
                        S.WriteLine("         ELSE")
                        S.WriteLine(MacroSubstSQLText("            UPDATE [:VALOR.ESQUEMA].SYS_DELETE SET MOMENTO = SYSDATE, USUARIO = USER, LOCAL = '[:VALOR.ESQUEMA]' WHERE TABELA = '" & tab("Tabela") & "' AND CHAVE || ''= " & ChavePrima & " || '';", Params))
                        S.WriteLine("         END IF;")
                        S.WriteLine("      ELSE")
                        S.WriteLine("         IF :NEW.SYS_STATUS = '+' THEN")
                        S.WriteLine("            :NEW.SYS_STATUS := 'I';")
                        S.WriteLine("         ELSIF :NEW.SYS_STATUS = '/' THEN")
                        S.WriteLine("            :NEW.SYS_STATUS := 'A';")
                        S.WriteLine("         ELSIF :NEW.SYS_STATUS = 'X' THEN")
                        S.WriteLine("            :NEW.SYS_STATUS := '';")
                        S.WriteLine("         ELSE")
                        S.WriteLine("            IF INSERTING THEN")
                        S.WriteLine("               :NEW.SYS_MOMENTO_CRIA := SYSDATE;")
                        S.WriteLine("               :NEW.SYS_USUARIO_CRIA := CONN_USER;")
                        S.WriteLine("               :NEW.SYS_LOCAL_CRIA := CONN_MACHINE || ' [' || CONN_IP || ']';")
                        S.WriteLine("               :NEW.SYS_MOMENTO_ATUALIZA := NULL;")
                        S.WriteLine("               :NEW.SYS_USUARIO_ATUALIZA := NULL;")
                        S.WriteLine("               :NEW.SYS_LOCAL_ATUALIZA := NULL;")
                        S.WriteLine("               :NEW.SYS_STATUS := 'I';")
                        S.WriteLine("            END IF;")
                        S.WriteLine("            IF UPDATING THEN")
                        S.WriteLine("               :NEW.SYS_MOMENTO_ATUALIZA := SYSDATE;")
                        S.WriteLine("               :NEW.SYS_USUARIO_ATUALIZA := CONN_USER;")
                        S.WriteLine("               :NEW.SYS_LOCAL_ATUALIZA := CONN_MACHINE || ' [' || CONN_IP || ']';")
                        S.WriteLine("               :NEW.SYS_STATUS := 'A';")
                        S.WriteLine("            END IF;")
                        S.WriteLine("         END IF;")
                        S.WriteLine("      END IF;")
                        S.WriteLine("   END IF;")
                        S.WriteLine("END;")
                        S.WriteLine("/")
                        S.WriteLine("")
                    End If
                Next

                ' atualização em cascata
                S.WriteLine("/* **********************************************************************************")
                S.WriteLine("   CRIAÇÃO DE TRIGGERS PARA ALTERAÇÃO EM CASCATA")
                S.WriteLine("*/")
                S.WriteLine("")
                For Each tab As Tabela In Me.ListaDeTabelas
                    Dim Texto As String = ""
                    ' para cada tabela, busca os relacionamentos pertinentes
                    ' será somente um trigger CCD (cascade) para cada tabela
                    ' neste trigger, todos os relacionamentos serão considerados
                    For Each rr As Rel In Me.ListaDeRels
                        If tab("Tabela") = rr("Tabela_1") Then
                            Dim Campos_1 As String() = Split(rr("Campo_1"), ";")
                            Dim Campos_N As String() = Split(rr("Campo_N"), ";")
                            Dim Crit As String = ""
                            Dim Def As String = ""
                            For z As Integer = 0 To Campos_1.Length - 1
                                Crit &= IIf(Crit <> "", " OR ", "") & "(:OLD." & Campos_1(z) & " <> :NEW." & Campos_1(z) & ")"
                                Def &= IIf(Def <> "", ", ", "") & Campos_N(z) & " = :NEW." & Campos_1(z)
                            Next
                            Dim Rels_Completo As String = ""
                            If Campos_1.Length > 1 Then
                                For z As Integer = 0 To Campos_1.Length - 1
                                    Rels_Completo &= IIf(Rels_Completo <> "", " AND ", "") & "NOT " & Campos_N(z) & " IS NULL"
                                Next
                            End If
                            Texto &= vbCrLf & "   --" & rr("Tabela_N") & vbCrLf
                            Texto &= "   IF " & Crit & " THEN" & vbCrLf
                            Texto &= MacroSubstSQLText("      UPDATE [:VALOR.ESQUEMA]." & rr("Tabela_N") & " SET " & Def & " WHERE " & Replace(Replace(Def, "NEW.", "OLD."), ", ", " AND ") & IIf(Rels_Completo <> "", " AND " & Rels_Completo, "") & ";", Params) & vbCrLf
                            Texto &= "   END IF;" & vbCrLf
                        End If
                    Next
                    If Texto <> "" Then
                        S.WriteLine(MacroSubstSQLText("CREATE OR REPLACE TRIGGER [:VALOR.ESQUEMA].AFT_" & tab("Tabela") & "_CCD", Params))
                        S.WriteLine(MacroSubstSQLText("AFTER UPDATE ON [:VALOR.ESQUEMA]." & tab("Tabela") & " FOR EACH ROW", Params))
                        S.WriteLine("BEGIN")
                        S.WriteLine(Texto)
                        S.WriteLine("END;")
                        S.WriteLine("/")
                        S.WriteLine("")
                    End If
                Next


                ' cria visões
                S.WriteLine("/* **********************************************************************************")
                S.WriteLine("   CRIAÇÃO DE VISÕES")
                S.WriteLine("*/")
                S.WriteLine("")
                For Each v As Visao In NZ(ListaDeVisoes, New ArrayList)
                    S.WriteLine(MacroSubstSQLText("CREATE OR REPLACE VIEW [:VALOR.ESQUEMA]." & v.Nome & " AS " & v.Texto, Params))
                    S.WriteLine("/")
                    S.WriteLine("")
                Next
                S.WriteLine("")

                ' cria objetos
                S.WriteLine("/* **********************************************************************************")
                S.WriteLine("   OUTROS CÓDIGOS EM GERAL")
                S.WriteLine("*/")
                S.WriteLine("")
                For Each ob As Obj In NZ(ListaDeObjs, New ArrayList)
                    S.WriteLine(MacroSubstSQLText(ob.Texto, Params))
                    S.WriteLine("/")
                    S.WriteLine("")
                Next

                ' cria usuários
                S.WriteLine("/* **********************************************************************************")
                S.WriteLine("   CRIAÇÃO DE USUÁRIOS")
                S.WriteLine("*/")
                S.WriteLine("")
                For Each us As Usuario In NZ(ListaDeUsuarios, New ArrayList)
                    S.WriteLine("CREATE USER " & us.Login & " IDENTIFIED BY " & Chr(34) & us.Senha & Chr(34))
                    S.WriteLine(MacroSubstSQLText("DEFAULT TABLESPACE T_[:ESQUEMA]_DAT", Params))
                    S.WriteLine("TEMPORARY TABLESPACE TEMP")
                    S.WriteLine("PROFILE DEFAULT")
                    S.WriteLine("ACCOUNT UNLOCK;")
                    For Each dr As Direito In us.ListaDeDireitos
                        S.WriteLine(MacroSubstSQLText("GRANT " & dr.Permissao & " ON [:ESQUEMA]." & dr.Objeto & " TO " & us.Login & ";", Params))
                    Next
                    S.WriteLine("")
                Next

                ' encerramento de script
                S.WriteLine("SPOOL OFF")
            End Sub

            ''' <summary>
            ''' Opera script a ser executado no oracle neste momento ou através de download do texto.
            ''' </summary>
            ''' <param name="ExecutaNoOracle">Executa direto no oracle.</param>
            ''' <param name="Parte">Parte(0=todas,1=sem restr e 2=restrições)</param>
            ''' <param name="ListaDeParams">Parâmetros como ESQUEMA,"ESQUEMA".</param>
            ''' <returns>Retorna texto de console da criação no oracle (caso executando através oracle).</returns>
            ''' <remarks></remarks>
            Public Function GravaOracle(ByVal ExecutaNoOracle As Boolean, ByVal Parte As GravaOracleParteTipo, ByVal ParamArray ListaDeParams() As Object) As String
                Dim Result As String = ""
                Dim Params As ArrayList = ParamArrayToArrayList(ListaDeParams)
                Dim TempArq As String = TemporaryFile()
                Dim S As New System.IO.StreamWriter(TempArq)

                If Parte <> GravaOracleParteTipo.Dois Then
                    GravaOracleSemRestr(S, ExecutaNoOracle, Params)
                End If
                If Parte <> GravaOracleParteTipo.Um Then
                    GravaOracleRestr(S, ExecutaNoOracle, Params)
                End If

                S.Flush()
                S.Close()
                If ExecutaNoOracle Then
                    ' executa script no oracle
                    Dim Psi As New System.Diagnostics.ProcessStartInfo("SQLPLUS.EXE", MacroSubstSQLText("[:VALOR.USUARIOSYS]/[:VALOR.SENHASYS]@[:VALOR.SERVICO]", Params))
                    Psi.UseShellExecute = False
                    Psi.RedirectStandardError = True
                    Psi.RedirectStandardInput = True
                    Psi.RedirectStandardOutput = True
                    Psi.WorkingDirectory = TemporaryDir()
                    Dim Proc As System.Diagnostics.Process = System.Diagnostics.Process.Start(Psi)
                    Dim StdIn As System.IO.StreamWriter = Proc.StandardInput
                    Dim StdOut As System.IO.StreamReader = Proc.StandardOutput
                    Dim StdErr As System.IO.StreamReader = Proc.StandardError
                    StdIn.WriteLine("@" & TempArq)
                    Proc.Close()
                    StdIn.Close()
                    Result = "Resultado:" & vbCrLf & StdOut.ReadToEnd & vbCrLf & vbCrLf & "Erro:" & vbCrLf & StdErr.ReadToEnd
                    StdOut.Close()
                    StdErr.Close()
                    Return Result
                End If

                Dim R As New IO.StreamReader(TempArq)
                Result = R.ReadToEnd()
                R.Close()
                Return Result
            End Function

            Public Sub GravaMySQL(ByVal Maquina As String, ByVal BancoDeDados As String, ByVal Usuario As String, ByVal Senha As String, ByVal ParamArray ListaDeParams() As Object)

                Dim Params As ArrayList = ParamArrayToArrayList(ListaDeParams)
                Dim Result As String = ""

                ' preparo do ambiente de criação
                Dim TempArq As String = TemporaryFile()
                Dim S As New System.IO.StreamWriter(TempArq)

                ' banco de dados

                S.WriteLine("DROP DATABASE IF EXISTS " & BancoDeDados & ";")
                S.WriteLine("CREATE DATABASE " & BancoDeDados & ";")

                S.WriteLine("USE " & BancoDeDados & ";")
                S.WriteLine("")

                ' cria tabelas
                For Each tab As Tabela In Me.ListaDeTabelas
                    If NZ(tab("Codigo"), "") <> "" Then

                        ' obtem chave primária
                        Dim ChavePrima As ArrayList = ParamArrayToArrayList(Split(tab("Chave_Prima"), ";"))

                        ' monta corpo dos campos
                        Dim Texto As String = ""
                        For Each camp As Campo In tab("Campos")

                            ' PROVISÓRIO VERIFICAR SE É ITEM DE NUMERAÇÃO AUTO
                            ' SERÁ COLOCADO EM TRIGGER
                            Dim AUTONUM As Boolean = False
                            If ChavePrima.Count = 1 AndAlso ChavePrima(0) = camp("campo") AndAlso Compare(TipoScriptToMySQL(camp), "INT (11)") Then
                                AUTONUM = True
                            End If

                            Texto &= IIf(Texto <> "", "," & Chr(13) & Chr(10) & "   ", "") & "`" & camp("Campo") & "` " & TipoScriptToMySQL(camp) & IIf(AUTONUM, " AUTO_INCREMENT", "")
                        Next

                        ' gera texto com campos padronizados
                        S.WriteLine("DROP TABLE IF EXISTS `" & tab.Tabela & "`;")
                        S.WriteLine("CREATE TABLE `" & tab.Tabela & "` (")
                        S.WriteLine("   " & Texto & ",")
                        S.WriteLine("   `SYS_MOMENTO_CRIA` DATETIME,")
                        S.WriteLine("   `SYS_USUARIO_CRIA` VARCHAR (100),")
                        S.WriteLine("   `SYS_LOCAL_CRIA` VARCHAR (100),")
                        S.WriteLine("   `SYS_MOMENTO_ATUALIZA` DATETIME,")
                        S.WriteLine("   `SYS_USUARIO_ATUALIZA` VARCHAR (100),")
                        S.WriteLine("   `SYS_LOCAL_ATUALIZA` VARCHAR (100),")
                        S.WriteLine("   `SYS_STATUS` CHAR (1)")
                        S.WriteLine("")

                        ' cria chave primária
                        If ChavePrima.Count > 0 Then
                            S.WriteLine(",  PRIMARY KEY  (`" & Join(ChavePrima.ToArray, "`, `") & "`)")
                        End If

                        S.WriteLine(") ENGINE=InnoDB DEFAULT CHARSET=latin1;")
                        S.WriteLine("")
                    End If
                Next

                If Exporta And Criterios.InfraSistema Then

                    ' cria tabelas de sistema
                    S.WriteLine("")
                    S.WriteLine("CREATE TABLE `sys_config_global` (")
                    S.WriteLine("   PARAM VARCHAR (100) PRIMARY KEY,")
                    S.WriteLine("   DESCR TEXT,")
                    S.WriteLine("   CONFIG TEXT")
                    S.WriteLine(") ENGINE=InnoDB DEFAULT CHARSET=latin1;")

                    ' S.WriteLine("")
                    ' S.WriteLine("ALTER TABLE `sys_config_global` COMMENT = 'Sistema | Armazena parâmetros globais do aplicativo.';")

                    S.WriteLine("")
                    S.WriteLine("CREATE TABLE `sys_config_usuario` (")
                    S.WriteLine("   USUARIO VARCHAR (100) PRIMARY KEY,")
                    S.WriteLine("   PARAM VARCHAR (100) PRIMARY KEY,")
                    S.WriteLine("   DESCR TEXT,")
                    S.WriteLine("   CONFIG TEXT")
                    S.WriteLine(") ENGINE=InnoDB DEFAULT CHARSET=latin1;")
                    S.WriteLine("")

                    S.WriteLine("CREATE TABLE `sys_delete` (")
                    S.WriteLine("   TABELA VARCHAR (100),")
                    S.WriteLine("   CHAVE VARCHAR (300),")
                    S.WriteLine("   MOMENTO DATETIME,")
                    S.WriteLine("   USUARIO VARCHAR (100),")
                    S.WriteLine("   LOCAL VARCHAR (100)")
                    S.WriteLine()
                    S.WriteLine(",  PRIMARY KEY (`TABELA`, `CHAVE`, `MOMENTO`)")
                    S.WriteLine(") ENGINE=InnoDB DEFAULT CHARSET=latin1;")
                    S.WriteLine("")
                    S.WriteLine("CREATE TABLE `sys_localid` (")
                    S.WriteLine("   NOME VARCHAR (20) PRIMARY KEY,")
                    S.WriteLine("   CORRENTE INT (6),")
                    S.WriteLine("   PACOTE INT (6),")
                    S.WriteLine("   PACOTE_REC INT (6),")
                    S.WriteLine("   MOMENTO DATETIME,")
                    S.WriteLine("   MOMENTO_REC DATETIME,")
                    S.WriteLine("   MODELO INT (6),")
                    S.WriteLine("   OBS VARCHAR (300)")
                    S.WriteLine(") ENGINE=InnoDB DEFAULT CHARSET=latin1;")
                    S.WriteLine("")
                    S.WriteLine("CREATE TABLE `sys_ocorrencia` (")
                    S.WriteLine("   SEQ INT (11) PRIMARY KEY AUTO_INCREMENT,")
                    S.WriteLine("   APLICACAO VARCHAR (100),")
                    S.WriteLine("   OCORRENCIA TEXT,")
                    S.WriteLine("   USUARIO VARCHAR (100),")
                    S.WriteLine("   MOMENTO DATETIME,")
                    S.WriteLine("   LOCAL VARCHAR (100)")
                    S.WriteLine(") ENGINE=InnoDB DEFAULT CHARSET=latin1;")
                    S.WriteLine("")
                End If

                ' cria os relacionamentos
                If Me.ListaDeRels.Count > 0 Then
                    S.WriteLine("")
                    For Each rr As Rel In Me.ListaDeRels
                        S.WriteLine("ALTER TABLE `" & rr("Tabela_N") & "` ADD CONSTRAINT `" & rr("Nome") & "` FOREIGN KEY(`" & Join(Split(rr("Campo_N"), ";"), "`,`") & "`) REFERENCES `" & rr("Tabela_1") & "` (`" & Join(Split(rr("Campo_1"), ";"), "`,`") & "`)" & IIf(rr("DELETE_CASCADE") <> "", " ON DELETE " & rr("DELETE_CASCADE"), "") & IIf(rr("UPDATE_CASCADE") <> "", " ON UPDATE " & rr("UPDATE_CASCADE"), "") & ";")
                    Next
                End If

                ' cria triggers
                S.WriteLine("DELIMITER ;;")
                For Each tab As Tabela In Me.ListaDeTabelas
                    If NZ(tab("Codigo"), "") <> "" Then

                        ' gera texto do trigger
                        S.WriteLine("")
                        S.WriteLine("DROP TRIGGER IF EXISTS BEF_" & tab("codigo").ToString.ToLower() & "_INS;;")
                        S.WriteLine("CREATE TRIGGER BEF_" & tab("CODIGO").ToString.ToLower() & "_INS BEFORE INSERT ON " & tab("TABELA").ToString.ToLower())
                        S.WriteLine("FOR EACH ROW")
                        S.WriteLine("BEGIN")

                        S.WriteLine("")
                        S.WriteLine("   DECLARE MACHINE TEXT;")
                        S.WriteLine("   SET MACHINE = IFNULL(@CONN_IP,'');")
                        S.WriteLine("   IF IFNULL(@CONN_MACHINE,'')<>'' THEN")
                        S.WriteLine("      IF @CONN_MACHINE<>MACHINE THEN")
                        S.WriteLine("         SET MACHINE = CONCAT(@CONN_MACHINE,'[',MACHINE,']');")
                        S.WriteLine("      END IF;")
                        S.WriteLine("   END IF;")
                        S.WriteLine("   IF @@HOSTNAME <> IFNULL(@CONN_IP,'') AND @@HOSTNAME <> IFNULL(@CONN_MACHINE,'') THEN")
                        S.WriteLine("      SET MACHINE = CONCAT(@@HOSTNAME, IF(MACHINE<>'',' (CLIENT:',''), MACHINE, IF(MACHINE<>'',')',''));")
                        S.WriteLine("   END IF;")
                        S.WriteLine("")

                        S.WriteLine("   SET NEW.SYS_MOMENTO_CRIA = NOW();")
                        S.WriteLine("   SET NEW.SYS_USUARIO_CRIA = IF(IFNULL(@CONN_USER,'')='',USER(),@CONN_USER);")
                        S.WriteLine("   SET NEW.SYS_LOCAL_CRIA = MACHINE;")
                        S.WriteLine("   SET NEW.SYS_MOMENTO_ATUALIZA = NULL;")
                        S.WriteLine("   SET NEW.SYS_USUARIO_ATUALIZA = NULL;")
                        S.WriteLine("   SET NEW.SYS_LOCAL_ATUALIZA = NULL;")
                        S.WriteLine("   SET NEW.SYS_STATUS = 'I';")
                        S.WriteLine("END;")
                        S.WriteLine(";;")

                        S.WriteLine("")
                        S.WriteLine("DROP TRIGGER IF EXISTS BEF_" & tab("codigo").ToString.ToLower() & "_UPD;;")
                        S.WriteLine("CREATE TRIGGER BEF_" & tab("CODIGO").ToString.ToLower() & "_UPD BEFORE UPDATE ON " & tab("TABELA").ToString.ToLower())
                        S.WriteLine("FOR EACH ROW")
                        S.WriteLine("BEGIN")

                        S.WriteLine("")
                        S.WriteLine("   DECLARE MACHINE TEXT;")
                        S.WriteLine("   SET MACHINE = IFNULL(@CONN_IP,'');")
                        S.WriteLine("   IF IFNULL(@CONN_MACHINE,'')<>'' THEN")
                        S.WriteLine("      IF @CONN_MACHINE<>MACHINE THEN")
                        S.WriteLine("         SET MACHINE = CONCAT(@CONN_MACHINE,'[',MACHINE,']');")
                        S.WriteLine("      END IF;")
                        S.WriteLine("   END IF;")
                        S.WriteLine("   IF @@HOSTNAME <> IFNULL(@CONN_IP,'') AND @@HOSTNAME <> IFNULL(@CONN_MACHINE,'') THEN")
                        S.WriteLine("      SET MACHINE = CONCAT(@@HOSTNAME, IF(MACHINE<>'',' (CLIENT:',''), MACHINE, IF(MACHINE<>'',')',''));")
                        S.WriteLine("   END IF;")
                        S.WriteLine("")

                        S.WriteLine("   SET NEW.SYS_MOMENTO_ATUALIZA = NOW();")
                        S.WriteLine("   SET NEW.SYS_USUARIO_ATUALIZA = IF(IFNULL(@CONN_USER,'')='',USER(),@CONN_USER);")
                        S.WriteLine("   SET NEW.SYS_LOCAL_ATUALIZA = MACHINE;")
                        S.WriteLine("   SET NEW.SYS_STATUS = 'A';")
                        S.WriteLine("END;")
                        S.WriteLine(";;")

                        Dim ChavePrima As String = Mid(Replace(";" & tab("Chave_Prima"), ";", " || OLD."), 5)

                        S.WriteLine("")
                        S.WriteLine("DROP TRIGGER IF EXISTS BEF_" & tab("codigo").ToString.ToLower() & "_DEL;;")
                        S.WriteLine("CREATE TRIGGER BEF_" & tab("CODIGO").ToString.ToLower() & "_DEL BEFORE DELETE ON " & tab("TABELA").ToString.ToLower())
                        S.WriteLine("FOR EACH ROW")
                        S.WriteLine("BEGIN")

                        S.WriteLine("")
                        S.WriteLine("   DECLARE TOT INTEGER;")
                        S.WriteLine("   DECLARE MACHINE TEXT;")
                        S.WriteLine("   SET MACHINE = IFNULL(@CONN_IP,'');")
                        S.WriteLine("   IF IFNULL(@CONN_MACHINE,'')<>'' THEN")
                        S.WriteLine("      IF @CONN_MACHINE<>MACHINE THEN")
                        S.WriteLine("         SET MACHINE = CONCAT(@CONN_MACHINE,'[',MACHINE,']');")
                        S.WriteLine("      END IF;")
                        S.WriteLine("   END IF;")
                        S.WriteLine("   IF @@HOSTNAME <> IFNULL(@CONN_IP,'') AND @@HOSTNAME <> IFNULL(@CONN_MACHINE,'') THEN")
                        S.WriteLine("      SET MACHINE = CONCAT(@@HOSTNAME, IF(MACHINE<>'',' (CLIENT:',''), MACHINE, IF(MACHINE<>'',')',''));")
                        S.WriteLine("   END IF;")
                        S.WriteLine("")

                        S.WriteLine("   SELECT COUNT(*) INTO TOT FROM sys_delete WHERE TABELA = '" & tab("TABELA").ToString.ToLower() & "' AND CHAVE = " & ChavePrima & ";")
                        S.WriteLine("   If TOT = 0 Then")
                        S.WriteLine("       INSERT INTO sys_delete SET")
                        S.WriteLine("       TABELA = '" & tab("TABELA").ToString.ToLower() & "',")
                        S.WriteLine("       CHAVE = " & ChavePrima & ",")
                        S.WriteLine("       MOMENTO = NOW(),")
                        S.WriteLine("       USUARIO = IF(IFNULL(@CONN_USER,'')='',USER(),@CONN_USER),")
                        S.WriteLine("       LOCAL = MACHINE;")
                        S.WriteLine("   Else")
                        S.WriteLine("       UPDATE sys_delete SET")
                        S.WriteLine("       TABELA = '" & tab("TABELA").ToString.ToLower() & "',")
                        S.WriteLine("       CHAVE = " & ChavePrima & ",")
                        S.WriteLine("       MOMENTO = NOW(),")
                        S.WriteLine("       USUARIO = IF(IFNULL(@CONN_USER,'')='',USER(),@CONN_USER),")
                        S.WriteLine("       LOCAL = MACHINE")
                        S.WriteLine("       WHERE TABELA = '" & tab("TABELA").ToString.ToLower() & "' AND CHAVE = " & ChavePrima & ";")
                        S.WriteLine("   END IF;")
                        S.WriteLine("END;")
                        S.WriteLine(";;")

                    End If
                Next
                S.WriteLine("DELIMITER ;")

                S.Close()
            End Sub
            Function Diferenca(ByVal Estrutura As Gerador, Optional ByVal page As Page = Nothing) As String
                Dim Result As String = ""

                ' sistema
                If NZ(Prop("chksistema", "checked", page), True) Then
                    If Me.Descr <> Estrutura.Descr Then
                        Result &= "Descr: " & NZV(Me.Descr, "<vazio>") & ComboSepDefault & NZV(Estrutura.Descr, "<vazio>") & vbCrLf
                    End If
                    If Me.Ver <> Estrutura.Ver Then
                        Result &= "Ver: " & NZV(Me.Ver, "<vazio>") & ComboSepDefault & NZV(Estrutura.Ver, "<vazio>") & vbCrLf
                    End If
                    If Me.Data_Ver <> Estrutura.Data_Ver Then
                        Result &= "Data Versão: " & NZV(Me.Data_Ver, "<vazio>") & ComboSepDefault & NZV(Estrutura.Data_Ver, "<vazio>") & vbCrLf
                    End If
                    If Me.Rev <> Estrutura.Rev Then
                        Result &= "Revisão: " & NZV(Me.Rev, "<vazio>") & ComboSepDefault & NZV(Estrutura.Rev, "<vazio>") & vbCrLf
                    End If
                End If

                Dim Segm As String = ""

                ' busca tabela que existe em x e não existe em y

                If NZ(Prop("chktab1s2", "checked", page), True) Then
                    Dim NomeTabsRels As ArrayList = ItemsToArrayList(Estrutura.ListaDeTabelas, "Tabela")
                    For Each TabX As Tabela In Me.ListaDeTabelas
                        If Not NomeTabsRels.Contains(TabX.Tabela) Then
                            Segm &= TabX.Tabela & ComboSepDefault & "<inexistente>" & vbCrLf
                        End If
                    Next
                End If

                ' tem em y e não tem em x tabela
                Dim NomeTabs As ArrayList = ItemsToArrayList(Me.ListaDeTabelas, "Tabela")

                If NZ(Prop("chktab2s1", "checked", page), True) Then
                    For Each TabY As Tabela In Estrutura.ListaDeTabelas
                        If Not NomeTabs.Contains(TabY.Tabela) Then
                            Segm &= "<inexistente>" & ComboSepDefault & TabY.Tabela & vbCrLf
                        End If
                    Next
                End If

                ' existe em x e y e são diferentes

                For Each TabY As Tabela In Estrutura.ListaDeTabelas
                    If NomeTabs.Contains(TabY.Tabela) Then
                        Dim TabX As Tabela = Tabelas(TabY.Tabela)
                        Dim Dif As String = TabX.Diferenca(TabY, page)
                        If Dif <> "" Then
                            Segm &= Dif & vbCrLf
                        End If
                    End If
                Next
                If Segm <> "" Then
                    Result &= "Tabelas:" & vbCrLf & InsereTab(Segm, Gerador_Tabula) & vbCrLf
                End If


                ' relacionamentos em base1 sem base2
                Segm = ""
                If NZ(Prop("chkrel1s2", "checked", page), True) Then
                    Dim NomeRels As ArrayList = ItemsToArrayList(Estrutura.ListaDeRels, "Nome")
                    For Each RelX As Rel In Me.ListaDeRels
                        If Not NomeRels.Contains(RelX.Nome) Then
                            Segm &= RelX.Nome & ComboSepDefault & "<inexistente>" & vbCrLf
                        End If
                    Next
                End If

                ' relacionamentos em base2 sem base1

                If NZ(Prop("chkrel2s1", "checked", page), True) Then
                    Dim NomeRels As ArrayList = ItemsToArrayList(Me.ListaDeRels, "Nome")
                    For Each RelY As Rel In Estrutura.ListaDeRels
                        If Not NomeRels.Contains(RelY.Nome) Then
                            Segm &= "<inexistente>" & ComboSepDefault & RelY.Nome & vbCrLf
                        End If
                    Next
                End If

                ' relacionamentos nas duas bases e diferentes

                If NZ(Prop("chkrel12", "checked", page), True) Then
                    Dim NomeRels As ArrayList = ItemsToArrayList(Me.ListaDeRels, "Nome")
                    For Each RelY As Rel In Estrutura.ListaDeRels
                        If NomeRels.Contains(RelY.Nome) Then
                            Dim RelX As Rel = Rels(RelY.Nome)
                            Dim Dif As String = RelX.Diferenca(RelY)
                            If Dif <> "" Then
                                Segm &= Dif & vbCrLf
                            End If
                        End If
                    Next
                End If
                If Segm <> "" Then
                    Result &= "Relacionamentos:" & vbCrLf & InsereTab(Segm, Gerador_Tabula) & vbCrLf
                End If


                ' concatena resultado
                Return Result
            End Function
            Function TipoOracleToScript(ByVal data_type As String, ByVal data_length As String, ByVal data_precision As String, ByVal data_scale As String) As String
                Return Microsoft.VisualBasic.Switch(data_type = "VARCHAR2", "VARCHAR2 (" & data_length & ")", data_type = "NUMBER", "NUMBER" & IIf(data_precision <> "", " (" & data_precision & IIf(data_scale <> "", "," & data_scale, "") & ")", ""), True, data_type)
            End Function
            Function TipoScriptToAccess(ByVal Campo As Gerador.Campo) As String
                Dim Tam As Integer = 0
                Dim Decim As Integer = 0

                If Campo("Tipo_Access") <> "" Then
                    Return Campo("Tipo_Access")
                End If
                If Campo("Tipo_Oracle") <> "" Then
                    Dim Tipo As MatchCollection = System.Text.RegularExpressions.Regex.Matches(Campo("Tipo_Oracle"), "\w+")
                    If Compare(Tipo(0).Value, "NUMBER") Then
                        If Tipo.Count > 1 Then
                            Tam = Tipo(1).Value
                        End If
                        If Tipo.Count > 2 Then
                            Decim = Val(Tipo(2).Value)
                        End If
                        If Tam = 1 And Decim = 0 Then
                            Return "BOOLEAN;1"
                        ElseIf Tam <= 6 And Decim = 0 Then
                            Return "INTEGER;3"
                        ElseIf Tam <= 10 And Decim = 0 Then
                            Return "LONG;4"
                        ElseIf Tam <= 12 And Decim = 0 Then
                            Return "DECIMAL;20"
                        ElseIf Tam <= 16 And Decim = 2 Then
                            Return "CURRENCY;5"
                        ElseIf Tam > 16 Or Decim > 2 Then
                            Return "DOUBLE;" & DAO_DataTypeEnum_dbDouble
                        Else
                            Throw New Exception("Tipo NUMÉRICO não previsto no campo " & Campo("Campo") & " em script to access, obtendo de oracle.")
                        End If
                    ElseIf Compare(Tipo(0).Value, "FLOAT") Then
                        Return "DOUBLE;" & DAO_DataTypeEnum_dbDouble
                    ElseIf Compare(Tipo(0).Value, "DATE") Then
                        Return "DATE;" & DAO_DataTypeEnum_dbDate
                    ElseIf Compare(Tipo(0).Value, "CLOB") Then
                        Return "MEMO;" & DAO_DataTypeEnum_dbMemo
                    ElseIf Compare(Tipo(0).Value, "BLOB") Then
                        Return "OBJETOOLE;" & DAO_DataTypeEnum_dbBinary
                    ElseIf Compare(Tipo(0).Value, "VARCHAR2") Then
                        If Tipo.Count > 1 Then
                            Tam = Tipo(1).Value
                        End If
                        If Tam < 255 Then
                            Return "TEXT;" & DAO_DataTypeEnum_dbText & " (" & Tipo(1).Value & ")"
                        Else
                            Return "MEMO;" & DAO_DataTypeEnum_dbMemo
                        End If
                    ElseIf Compare(Tipo(0).Value, "CHAR") Then
                        Return "TEXT;" & DAO_DataTypeEnum_dbText & " (1)"
                    End If
                End If


                If Campo("Tipo_MySql") <> "" Then
                    Dim Tipo As MatchCollection = System.Text.RegularExpressions.Regex.Matches(Campo("Tipo_MySQL"), "\w+")
                    If InStr(";BIT;TINYINT;SMALLINT;MEDIUMINT;INT;BITINT;FLOAT;DOUBLE;DECIMAL;DEC;", ";" & UCase(Tipo(0).Value) & ";") <> 0 Then
                        If Tipo.Count > 1 Then
                            Tam = Tipo(1).Value
                        End If
                        If Tipo.Count > 2 Then
                            Decim = Val(Tipo(2).Value)
                        End If
                        If Tam <= 6 And Decim = 0 Then
                            Return "INTEGER;3"
                        ElseIf Tam <= 10 And Decim = 0 Then
                            Return "LONG;4"
                        ElseIf Tam <= 12 And Decim = 0 Then
                            Return "DECIMAL;20"
                        ElseIf Tam <= 16 And Decim = 2 Then
                            Return "CURRENCY;6"
                        ElseIf Tam <= 16 Then
                            Return "DOUBLE;5"
                        Else
                            Throw New Exception("Tipo NUMÉRICO não previsto no campo " & Campo("Campo") & " em script to access, obtendo de oracle.")
                        End If
                    ElseIf InStr(";BOOL;BOOLEAN;", ";" & UCase(Tipo(0).Value) & ";") <> 0 Then
                        Return "BOOLEAN;" & DAO_DataTypeEnum_dbBoolean
                    ElseIf InStr(";DATE;DATETIME;TIMESTAMP;TIME;YEAR;", ";" & UCase(Tipo(0).Value) & ";") <> 0 Then
                        Return "DATE;" & DAO_DataTypeEnum_dbDate
                    ElseIf InStr(";BINARY;VARBINARY;TINYBLOB;BLOB;MEDIUMBLOB;LONGBLOB;", ";" & UCase(Tipo(0).Value) & ";") <> 0 Then
                        Return "OBJETOOLE;" & DAO_DataTypeEnum_dbBinary
                    ElseIf InStr(";TEXT;LONGTEXT;", ";" & UCase(Tipo(0).Value) & ";") <> 0 Then
                        Return "MEMO;" & DAO_DataTypeEnum_dbMemo
                    ElseIf InStr(";CHAR;VARCHAR;TINYTEXT;MEDIUMTEXT;TEXT;LONGTEXT;", ";" & UCase(Tipo(0).Value) & ";") <> 0 Then
                        If Tipo.Count > 1 Then
                            Tam = Tipo(1).Value
                        End If
                        If Tam < 255 Then
                            Return "TEXT;" & DAO_DataTypeEnum_dbText & " (" & Tipo(1).Value & ")"
                        Else
                            Return "MEMO;" & DAO_DataTypeEnum_dbMemo
                        End If
                    End If
                End If
                Throw New Exception("Tipo não previsto no campo " & Campo("Campo") & " em script to access, obtendo de oracle.")
            End Function
            Function TipoScriptToOracle(ByVal Campo As Gerador.Campo) As String
                If Campo("Tipo_Oracle") <> "" Then
                    Return Campo("Tipo_Oracle")
                End If
                If Campo("Tipo_Access") <> "" Then
                    Dim Tipo As Match = RegexMatches(Campo("Tipo_Access"), "(.*);([0-9]*)($|[ ]*\(([0-9]*)\))")
                    If Compare(Tipo.Groups(1).Value, "LONG") Then
                        Return "NUMBER (11,0)"
                    ElseIf Compare(Tipo.Groups(1).Value, "BOOLEAN") Then
                        Return "NUMBER (1,0)"
                    ElseIf Compare(Tipo.Groups(1).Value, "CURRENCY") Then
                        Return "NUMBER (16,2)"
                    ElseIf Compare(Tipo.Groups(1).Value, "DATE") Then
                        Return "DATE"
                    ElseIf Compare(Tipo.Groups(1).Value, "DECIMAL") Then
                        Return "NUMBER (12,0)"
                    ElseIf Compare(Tipo.Groups(1).Value, "DOUBLE") Then
                        Return "NUMBER (16,12)"
                    ElseIf Compare(Tipo.Groups(1).Value, "SINGLE") Then
                        Return "NUMBER (8,6)"
                    ElseIf Compare(Tipo.Groups(1).Value, "INTEGER") Then
                        Return "NUMBER (6,0)"
                    ElseIf Compare(Tipo.Groups(1).Value, "MEMO") Then
                        Return "CLOB"
                    ElseIf Compare(Tipo.Groups(1).Value, "TEXT") Then
                        Return "VARCHAR2 (" & Tipo.Groups(4).Value & ")"
                    ElseIf Compare(Tipo.Groups(1).Value, "OBJETOOLE") Then
                        Return "BLOB"
                    Else
                        Throw New Exception("Tipo não previsto no campo " & Campo("Campo") & " em script to oracle, obtendo de Access (" & Tipo.Groups(1).Value & ").")
                        Exit Function
                    End If
                End If
                Throw New Exception("Tipo vazio em campo " & Campo("campo") & ".")
            End Function
            Function TipoScriptToMySQL(ByVal Campo As Gerador.Campo) As String
                If Campo("Tipo_MySQL") <> "" Then
                    Return Campo("Tipo_MySQL")
                End If
                If Campo("Tipo_Access") <> "" Then
                    Dim Tipo As Match = RegexMatches(Campo("Tipo_Access"), "(.*);([0-9]*)($|[ ]*\(([0-9]*)\))")
                    If Compare(Tipo.Groups(1).Value, "LONG") Then
                        Return "INT (11)"
                    ElseIf Compare(Tipo.Groups(1).Value, "BOOLEAN") Then
                        Return "INT (1)"
                    ElseIf Compare(Tipo.Groups(1).Value, "CURRENCY") Then
                        Return "DECIMAL (16,2)"
                    ElseIf Compare(Tipo.Groups(1).Value, "DATE") Then
                        Return "DATETIME"
                    ElseIf Compare(Tipo.Groups(1).Value, "DECIMAL") Then
                        Return "INT (12)"
                    ElseIf Compare(Tipo.Groups(1).Value, "DOUBLE") Then
                        Return "FLOAT"
                    ElseIf Compare(Tipo.Groups(1).Value, "INTEGER") Then
                        Return "INT (6)"
                    ElseIf Compare(Tipo.Groups(1).Value, "MEMO") Then
                        Return "TEXT"
                    ElseIf Compare(Tipo.Groups(1).Value, "TEXT") Then
                        Return "VARCHAR (" & Tipo.Groups(4).Value & ")"
                    ElseIf Compare(Tipo, "OBJETOOLE") Then
                        Return "BLOB"
                    Else
                        Throw New Exception("Tipo não previsto no campo " & Campo("Campo") & " em script to oracle, obtendo de Access (" & Tipo.Groups(1).Value & ").")
                        Exit Function
                    End If
                End If
                Throw New Exception("Tipo vazio ou pre-definição de ORACLE não concluída para MYSQL " & Campo("campo") & ".")
            End Function
            Shared ReadOnly Property DefsCampo(ByVal StrGerador As String, ByVal Sistema As String, ByVal Tabela As String, ByVal NomeCampo As String) As DataRow
                Get
                    Dim Row As DataRow = Nothing
                    Try
                        Dim DS As DataSet = DSCarrega("SELECT * FROM GER_CAMPO WHERE SISTEMA=:SISTEMA AND TABELA=:TABELA AND CAMPO=:CAMPO", StrGerador, ":SISTEMA", Sistema, ":TABELA", Tabela, ":CAMPO", NomeCampo)
                        With DS.Tables(0).Rows(0)
                            If NZ(.Item("PROP_EXTEND"), "") <> "" Then
                                Dim Props As New ElementosStr(.Item("PROP_EXTEND"), vbCrLf, ":")
                                For Each Prop As ElementoStr In Props.Elementos
                                    DS.Tables(0).Columns.Add(Prop.Nome)
                                    .Item(Prop.Nome) = Prop.Conteudo
                                Next
                            End If
                        End With
                        Row = DS.Tables(0).Rows(0)
                    Catch
                    End Try
                    Return Row
                End Get
            End Property
            Shared ReadOnly Property DefsProp(ByVal Def As DataRow, ByVal NomeProp As String) As Object
                Get
                    Try
                        Return Def(NomeProp)
                    Catch
                    End Try
                    Return Nothing
                End Get
            End Property

            Sub GaranteExistenciaDasListas(Optional ByVal Limpando As Boolean = False)
                If IsNothing(ListaDeTabelas) Or Limpando Then
                    ListaDeTabelas = New List(Of Tabela)
                End If
                If IsNothing(ListaDeRels) Or Limpando Then
                    ListaDeRels = New List(Of Rel)
                End If
                If IsNothing(ListaDeClasses) Or Limpando Then
                    ListaDeClasses = New List(Of Classe)
                End If
                If IsNothing(ListaDeVisoes) Or Limpando Then
                    ListaDeVisoes = New List(Of Visao)
                End If
                If IsNothing(ListaDeObjs) Or Limpando Then
                    ListaDeObjs = New List(Of Obj)
                End If
                If IsNothing(ListaDeUsuarios) Or Limpando Then
                    ListaDeUsuarios = New List(Of Usuario)
                End If
                If IsNothing(ListaDeDireitos) Or Limpando Then
                    ListaDeDireitos = New List(Of Direito)
                End If
            End Sub

            Public Sub CarregaXML(ByVal TextoOuArquivo As String)
                If TextoOuArquivo.StartsWith("<NewDataSet>") Then
                    XML = TextoOuArquivo
                    Exit Sub
                End If
                XML = New System.IO.StreamReader(FileExpr(TextoOuArquivo), True).ReadToEnd
            End Sub

            Public Property XML() As String
                Get

                    Dim DS As New System.Data.DataSet
                    If Exporta And Criterios.Sistema Then
                        DS.Tables.Add(TipoComoTabela("Geral", Me, "(?is)^(?!lista)"))
                    End If

                    If Exporta And Criterios.Tabela Then
                        DS.Tables.Add(TipoComoTabela("Tabela", ListaDeTabelas, "(?is)^(?!lista)"))
                        If Not Exporta And Criterios.Comentario Then
                            DS.Tables("Tabela").Columns.Remove("Descr")
                            DS.Tables("Tabela").Columns.Remove("Etiq")
                        End If
                    End If

                    If Exporta And Criterios.Campo Then
                        DS.Tables.Add(TipoComoTabela("Campo", ListaDeCampos, "(?is)^(?!lista)"))
                        If Not Exporta And Criterios.Comentario Then
                            DS.Tables("Campo").Columns.Remove("Descr")
                            DS.Tables("Campo").Columns.Remove("Etiq")
                        End If
                    End If

                    If Exporta And Criterios.Relacionamento Then
                        DS.Tables.Add(TipoComoTabela("Rels", ListaDeRels, "(?is)^(?!lista)"))
                    End If

                    If Exporta And Criterios.Classe Then
                        DS.Tables.Add(TipoComoTabela("Classe", ListaDeClasses, "(?is)^(?!lista)"))
                    End If

                    If Exporta And Criterios.Visao Then
                        DS.Tables.Add(TipoComoTabela("Visao", ListaDeVisoes, "(?is)^(?!lista)"))
                    End If

                    If Exporta And Criterios.Objeto Then
                        DS.Tables.Add(TipoComoTabela("Obj", ListaDeObjs, "(?is)^(?!lista)"))
                    End If

                    If Exporta And Criterios.Usuario Then
                        DS.Tables.Add(TipoComoTabela("Usuario", ListaDeUsuarios, "(?is)^(?!lista)"))
                    End If

                    If Exporta And Criterios.Direito Then
                        DS.Tables.Add(TipoComoTabela("Direito", ListaDeDireitos, "(?is)^(?!lista)"))
                    End If

                    Return "<NewDataSet>" & vbCrLf & DS.GetXmlSchema.Replace("<?xml version=""1.0"" encoding=""utf-16""?>", "") & vbCrLf & New Icraft.IcftBase.RegexHtml(DS.GetXml, "NewDataSet").Inner & vbCrLf & "</NewDataSet>"
                End Get
                Set(ByVal value As String)
                    Dim DS As New System.Data.DataSet

                    DS.ReadXml(TextoEmStream(value), XmlReadMode.Auto)

                    GaranteExistenciaDasListas(Importa And Criterios.Iniciar)

                    For Each TB As DataTable In DS.Tables
                        For Each Linha As DataRow In TB.Rows
                            Select Case TB.TableName
                                Case "Geral"
                                    If Importa And Criterios.Sistema Then
                                        Sistema = NZ(Linha("Sistema"), "")
                                        Descr = NZ(Linha("Descr"), "")
                                        Ver = NZ(Linha("Ver"), "")
                                        Data_Ver = NZ(Linha("Data_Ver"), "")
                                        Rev = NZ(Linha("Rev"), "")
                                    End If
                                Case "Tabela"
                                    If Importa And Criterios.Tabela Then
                                        ListaDeTabelas.Add(New Tabela(NZ(Linha("Tabela"), ""), NZ(Linha("Ordem"), 0), NZ(Linha("Chave_Prima"), ""), Nothing, NZ(Linha("Codigo"), ""), NZ(Linha("Classe"), ""), "", ""))
                                    End If

                                    If Importa And Criterios.Comentario Then
                                        Tabelas(NZ(Linha("Tabela"), "")).Etiq = NZ(Linha("Etiq"), "")
                                        Tabelas(NZ(Linha("Tabela"), "")).Descr = NZ(Linha("Descr"), "")
                                    End If

                                Case "Campo"
                                    If Importa And Criterios.Campo Then
                                        If IsNothing(Tabelas(NZ(Linha("Tabela"), "")).ListaDeCampos) Then
                                            Tabelas(NZ(Linha("Tabela"), "")).ListaDeCampos = New List(Of Campo)
                                        End If
                                        Tabelas(NZ(Linha("Tabela"), "")).ListaDeCampos.Add(New Campo(NZ(Linha("Tabela"), ""), NZ(Linha("Campo"), ""), NZ(Linha("Ordem"), ""), NZ(Linha("Tipo_Access"), ""), NZ(Linha("Tipo_Oracle"), ""), NZ(Linha("Tipo_MySQL"), ""), "", "", NZ(Linha("Prop_Extend"), "")))
                                    End If

                                    If Importa And Criterios.Comentario Then
                                        Tabelas(NZ(Linha("Tabela"), "")).Campos(NZ(Linha("Campo"), "")).Etiq = NZ(Linha("Etiq"), "")
                                        Tabelas(NZ(Linha("Tabela"), "")).Campos(NZ(Linha("Campo"), "")).Descr = NZ(Linha("Descr"), "")
                                    End If

                                Case "Rels"
                                    If Importa And Criterios.Relacionamento Then
                                        ListaDeRels.Add(New Rel(NZ(Linha("Nome"), ""), NZ(Linha("Tabela_1"), ""), NZ(Linha("Campo_1"), ""), NZ(Linha("Tabela_N"), ""), NZ(Linha("Campo_N"), ""), NZ(Linha("Delete_Cascade"), ""), NZ(Linha("Update_Cascade"), ""), NZ(Linha("Obrig"), "")))
                                    End If
                                Case "Classe"
                                    If Importa And Criterios.Classe Then
                                        ListaDeClasses.Add(New Classe(NZ(Linha("Classe"), ""), NZ(Linha("Descr"), "")))
                                    End If
                                Case "Visao"
                                    If Importa And Criterios.Visao Then
                                        ListaDeVisoes.Add(New Visao(NZ(Linha("Nome"), ""), NZ(Linha("Classe"), ""), NZ(Linha("Texto"), "")))
                                    End If
                                Case "Obj"
                                    If Importa And Criterios.Objeto Then
                                        ListaDeObjs.Add(New Obj(NZ(Linha("Tipo"), ""), NZ(Linha("Ordem"), ""), NZ(Linha("Texto"), ""), NZ(Linha("Descr"), "")))
                                    End If
                                Case "Usuario"
                                    If Importa And Criterios.Usuario Then
                                        ListaDeUsuarios.Add(New Usuario(NZ(Linha("Login"), ""), NZ(Linha("Grupo"), ""), NZ(Linha("Senha"), ""), NZ(Linha("Nome"), ""), NZ(Linha("Depto"), ""), NZ(Linha("Obs"), "")))
                                    End If
                                Case "Direito"
                                    If Importa And Criterios.Direito Then
                                        ListaDeDireitos.Add(New Direito(NZ(Linha("Tipo"), ""), NZ(Linha("Objeto"), ""), NZ(Linha("Usuario"), ""), NZ(Linha("Permissao"), "")))
                                    End If
                            End Select
                        Next
                    Next
                End Set
            End Property
        End Class


        ''' <summary>
        ''' Transforma um objeto com seus atributos em uma tabela.
        ''' </summary>
        ''' <param name="Objeto">Objeto contendo os atributos.</param>
        ''' <param name="Filtro">Regex contendo termos de seleção dos atributos ex.: não começado por... (?is)^(?!teste).</param>
        ''' <returns>Datatable contendo os atributos.</returns>
        ''' <remarks></remarks>
        Public Shared Function TipoComoTabela(ByVal NomeTab As String, ByVal Objeto As Object, Optional ByVal Filtro As String = "", Optional ByVal TB As DataTable = Nothing) As DataTable
            If IsNothing(TB) Then
                TB = New DataTable(NomeTab)
            End If

            Try
                For Each Linha As Object In Objeto
                    If Not IsNothing(Linha) Then
                        TipoComoTabela(NomeTab, Linha, Filtro, TB)
                    End If
                Next
                Return TB
            Catch ex As Exception
                TB = TB
            End Try

            If Not IsNothing(Objeto) Then
                Dim Vals As New ArrayList
                Dim Estrut As Boolean = TB.Columns.Count <> 0
                For Each Item As System.Reflection.FieldInfo In Objeto.GetType.GetFields
                    If Filtro = "" OrElse Regex.Match(Item.Name, Filtro).Success Then
                        If Not Estrut Then
                            TB.Columns.Add(Item.Name)
                        End If
                        Vals.Add(Item.GetValue(Objeto))
                    End If
                Next
                TB.Rows.Add(Vals.ToArray)
            End If
            Return TB
        End Function

        ''' <summary>
        ''' Transforma texto em stream.
        ''' </summary>
        ''' <param name="Texto">Texto a ser colocado em stream.</param>
        ''' <returns>Stream.</returns>
        ''' <remarks></remarks>
        Public Shared Function TextoEmStream(ByVal Texto As String) As Stream
            Dim Memo As New IO.MemoryStream
            Dim MemoGrava As New IO.StreamWriter(Memo)
            MemoGrava.Write(Texto)
            MemoGrava.Flush()
            Memo.Position = 0
            Return Memo
        End Function






        Public Class TNSNamesReader

            Private strOracleHome As String = ""
            Private strTNSNAMESORAFilePath As String = ""


            Public Function GetOracleHome() As String
                Dim rgkActualHome As RegistryKey = Nothing
                Dim strOraHome As String = ""

                Dim rgkLM As RegistryKey = Registry.LocalMachine
                Dim rgkAllHome As RegistryKey = rgkLM.OpenSubKey("SOFTWARE\ORACLE\ALL_HOMES")

                If Not IsNothing(rgkAllHome) Then
                    Dim objLastHome As Object = rgkAllHome.GetValue("LAST_HOME")
                    Dim strLastHome = ""
                    strLastHome = objLastHome.ToString()
                    rgkActualHome = Registry.LocalMachine.OpenSubKey("SOFTWARE\ORACLE\HOME" + strLastHome)
                Else
                    For Each Reg As String In rgkLM.OpenSubKey("SOFTWARE\ORACLE").GetSubKeyNames.Reverse.ToArray
                        If Not IsNothing(Registry.LocalMachine.OpenSubKey("SOFTWARE\ORACLE\" + Reg)) Then
                            rgkActualHome = Registry.LocalMachine.OpenSubKey("SOFTWARE\ORACLE\" + Reg)
                            If Not IsNothing(rgkActualHome.GetValue("ORACLE_HOME")) Then
                                Exit For
                            End If
                        End If
                    Next
                End If

                If Not IsNothing(rgkActualHome) Then
                    Dim objOraHome As Object = rgkActualHome.GetValue("ORACLE_HOME")
                    strOraHome = objOraHome.ToString()
                    strOracleHome = strOraHome
                    Return strOraHome
                End If
                Return ""
            End Function

            Public Function GetTNSNAMESORAFilePath() As String
                If Not Me.GetOracleHome.Equals("") Then
                    strTNSNAMESORAFilePath = strOracleHome + "\NETWORK\ADMIN\TNSNAMES.ORA"
                    If Not (System.IO.File.Exists(strTNSNAMESORAFilePath)) Then
                        strTNSNAMESORAFilePath = strOracleHome + "\NET80\ADMIN\TNSNAMES.ORA"
                        Return strTNSNAMESORAFilePath
                    Else
                        Return strTNSNAMESORAFilePath
                    End If
                Else
                    Return ""
                End If
            End Function


            Public Function LoadTNSNames() As Collection
                Dim DBNamesCollection As New Collection
                Dim RegExPattern As String = "[\n][\s]*[^\(][a-zA-Z0-9_.]+[\s]*=[\s]*\("

                GetTNSNAMESORAFilePath()
                DBNamesCollection.Add("")
                If Not strTNSNAMESORAFilePath.Equals("") Then
                    Try
                        Dim fiTNS As New System.IO.FileInfo(strTNSNAMESORAFilePath)
                        If (fiTNS.Exists) Then
                            If (fiTNS.Length > 0) Then
                                Dim iCount As Integer
                                Try
                                    For iCount = 0 To Regex.Matches(My.Computer.FileSystem.ReadAllText(fiTNS.FullName), RegExPattern).Count - 1
                                        DBNamesCollection.Add(Regex.Matches(My.Computer.FileSystem.ReadAllText(fiTNS.FullName), RegExPattern).Item(iCount).Value.Trim.Substring(0, Regex.Matches(My.Computer.FileSystem.ReadAllText(fiTNS.FullName), RegExPattern).Item(iCount).Value.Trim.IndexOf(" ")))
                                    Next
                                Catch ex As Exception
                                    Throw New Exception(ex.Message)
                                End Try
                            End If
                        End If
                    Catch ex As Exception
                        Throw New Exception(ex.Message)
                    End Try
                End If
                Return DBNamesCollection
            End Function


        End Class



        Public Shared Function Soma(ByVal ParamArray Valores() As Object) As Double
            Dim Lista As ArrayList = ParamArrayToArrayList(Valores)
            Dim Total As Double = 0
            For Each V As Double In Lista
                Total += V
            Next
            Return Total
        End Function



        Public Shared Function FiltroCampoConteudo(ByVal ParamArray Definicoes() As Object) As String
            Dim Conteudo As ArrayList = ParamArrayToArrayList(Definicoes)
            Dim Result As New StringBuilder
            Dim z As Integer = 0
            Do While z < Conteudo.Count
                If Result.Length > 0 Then
                    Result.Append(" and ")
                End If
                Result.Append(Conteudo(z).ToString)
                z += 1
                Result.Append("=")
                Result.Append(SqlExpr(Conteudo(z)))
                z += 1
            Loop
            Return Result.ToString
        End Function


        Public Shared Function TipoAccessToScript(ByVal fld As Object) As String
            Dim ret As String
            If fld.Type = 1 Then
                ' boolean
                ret = "BOOLEAN;1"
            ElseIf fld.Type = 5 Then
                ' currency
                ret = "CURRENCY;5"
            ElseIf fld.Type = 8 Then
                ' date
                ret = "DATE;8"
            ElseIf fld.Type = 20 Then
                ' decimal
                ret = "DECIMAL;20"
            ElseIf fld.Type = 6 Then
                ' single
                ret = "SINGLE;6"
            ElseIf fld.Type = 7 Then
                ' double
                ret = "DOUBLE;7"
            ElseIf fld.Type = 3 Then
                ' integer
                ret = "INTEGER;3"
            ElseIf fld.Type = 4 Then
                ' long
                ret = "LONG;4"
            ElseIf fld.Type = 12 Then
                ' memo
                ret = "MEMO;12"
            ElseIf fld.Type = 10 Then
                ' text
                ret = "TEXT;10 (" & fld.Size & ")"
            ElseIf fld.Type = 11 Then
                ' objeto ole
                ret = "OBJETOOLE;11"
            ElseIf fld.Type = 9 Then
                ' binário
                ret = "BINARY;9"
            Else
                Throw New Exception("TipoAccess não previsto:" & fld.Type & ".")
            End If
            Return ret
        End Function

        Class Predicados
            Private _valor As Object
            Sub New(ByVal Valor As Object)
                _valor = Valor
            End Sub
            Public Function Exists(ByVal ValorVetor As Object) As Boolean
                Try
                    Return ValorVetor = _valor
                Catch
                End Try
                Return False
            End Function
        End Class


#If _MyType = "Web" Then
        Public Shared Sub RegistraControleComoPostBack(ByVal Page As Page, ByVal Ctl As Object)
            Dim Script As ScriptManager = Icraft.IcftBase.Form.BuscaPrimeiroTipo(Page, GetType(ScriptManager))
            If Not IsNothing(Script) Then
                Script.RegisterPostBackControl(Ctl)
            End If
        End Sub
#End If


        Public Shared Function HTMLD(ByVal Texto As String) As String
            Dim Ret As String = NZ(Texto, "")
            Ret = TiraHtml(Ret)
            Return System.Web.HttpUtility.HtmlDecode(Ret)
        End Function

        Public Shared Function TiraHtml(ByVal Texto As String) As String
            Texto = Regex.Replace(Texto, "(?is)<br[ ]*/>", vbCrLf)
            Texto = Regex.Replace(Texto, "(?is)<br[ ]*></br>", vbCrLf)
            Texto = Regex.Replace(Texto, "<.*?>", "")
            Texto = Regex.Replace(Texto, "</.*?>", "")
            Return Texto
        End Function

        Public Shared Function ItemEncode(ByVal Texto As String) As String
            Texto = NZ(Texto, "")
            Texto = Texto.Replace("&", "&#38&")
            Texto = Texto.Replace("_", "&#95&")
            Texto = Texto.Replace(";", "&#59&")
            Texto = Texto.Replace("""", "&#34&")
            Texto = Texto.Replace("|", "&#124&")
            Texto = Texto.Replace(":", "&#58&")
            Return Texto.Replace("&", "_")
        End Function

        Public Shared Function ItemDecode(ByVal texto As String) As String
            texto = Regex.Replace(texto, "_(#[0-9]{1,3})_", "&$1;")
            Return HttpUtility.HtmlDecode(texto)
        End Function

        Public Shared Sub RecuperaControles(ByVal Container As Object, ByVal Conteudo As String)
            For Each Item As String In Split(Conteudo, ";")
                If Trim(Item) <> "" Then
                    Dim ItemD() As String = Split(Item, ":")
                    Dim CampoId As String = ItemDecode(ItemD(0))
                    Dim CampoConteudo As String = ItemDecode(ItemD(1))
                    Dim Ctl As Object = Icraft.IcftBase.Form.FindGeral(Container, CampoId)
                    If Not IsNothing(Ctl) Then
                        Prop(Ctl) = CampoConteudo
                    End If
                End If
            Next
        End Sub

        Public Shared Function SalvaControles(ByVal Container As Object, ByVal Prefix As String, ByRef Ignorar As String) As String
            Dim Ret As New StringBuilder

            Try
                Dim id As String = Container.id
                If Not TemNaLista(Ignorar, Container.UniqueId) Then
                    Ignorar &= Container.uniqueid & ";"
                    If Prefix = "" OrElse id.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase) Then
                        Ret.Append(ItemEncode(id))
                        Ret.Append(":" & ItemEncode(NZV(Prop(Container), "")))
                        Ret.Append(";")
                    End If
                End If
            Catch
            End Try

            Dim Cont As Object = Container
            Try
                Dim QtdControls As Integer = Container.Controls.Count()
                If QtdControls > 0 Then
                    Cont = Container.Controls
                End If
            Catch
            End Try

            Try
                For Each Ctl As Object In Cont
                    Dim ContSub As Object = Ctl
                    Try
                        Dim QtdControls As Integer = Ctl.Controls.count
                        If QtdControls > 0 Then
                            ContSub = Ctl.Controls
                        End If
                    Catch
                    End Try
                    For Each SubCtl As Object In ContSub
                        Ret.Append(SalvaControles(SubCtl, Prefix, Ignorar))
                    Next
                Next
            Catch
            End Try
            Return Ret.ToString
        End Function

        Shared Sub RecuperaDoForm(ByVal Page As Page, ByVal Pref As String)
            For Each Item As String In Page.Request.Form
                If Item.StartsWith(Pref) Then
                    Prop(Page.FindControl(Item)) = Page.Request.Form(Item)
                End If
            Next
        End Sub

        Shared Function ObtemCor(ByVal Texto As String) As Color
            Try
                If Texto.StartsWith("#") Then
                    Dim H As String = Mid(Texto, 2)
                    H = Microsoft.VisualBasic.Left("FF000000", 8 - Len(H)) & H
                    Return Color.FromArgb(Val("&h" & H))
                Else
                    Return Color.FromName(Texto)
                End If
            Catch
            End Try
            Return Color.Transparent
        End Function



        Shared Property RegAplKey(ByVal NomeApl As String, ByVal Param As String, Optional ByVal Empresa As String = "Intercraft")
            Get
                Return RegMachineKey(ExprExpr("\", "/", "SOFTWARE", Empresa, NomeApl), Param)
            End Get
            Set(ByVal value)
                RegMachineKey(ExprExpr("\", "/", "SOFTWARE", Empresa, NomeApl), Param) = value
            End Set
        End Property

        Shared Property RegMachineKey(ByVal Local As String, ByVal Param As String)
            Get
                Dim Reg As RegistryKey
                Dim Ret As Object = Nothing
                Reg = Registry.LocalMachine.OpenSubKey(Local)
                If Not IsNothing(Reg) Then
                    Ret = Reg.GetValue(Param)
                End If
                Return Ret
            End Get
            Set(ByVal value)
                Dim Reg As RegistryKey = Registry.LocalMachine
                Dim RegNovo As RegistryKey
                RegNovo = Registry.LocalMachine.OpenSubKey(Local, True)
                If IsNothing(RegNovo) Then
                    RegNovo = Reg.CreateSubKey(Local)
                End If
                If Not IsNothing(RegNovo) Then
                    RegNovo.SetValue(Param, value)
                    RegNovo.Close()
                End If
                Reg.Close()
                RegNovo.Close()
            End Set
        End Property

        Class DirReplica
            Sub New(ByVal DirOrigem As String, ByVal DirDestino As String, ByVal IncluirSubDir As Boolean, Optional ByVal ApagarQuandoEncontrar As String = "")
                Me.DirOrigem = DirOrigem
                Me.DirDestino = DirDestino
                Me.IncluiSub = IncluirSubDir
                Me._ApagarQuandoEncontrar = ApagarQuandoEncontrar
            End Sub
            Private _ApagarQuandoEncontrar As String
            Private _dirorigem As String
            Public Property DirOrigem() As String
                Get
                    Return _dirorigem
                End Get
                Set(ByVal value As String)
                    _dirorigem = value
                End Set
            End Property
            Private _dirdestino As String
            Public Property DirDestino() As String
                Get
                    Return _dirdestino
                End Get
                Set(ByVal value As String)
                    _dirdestino = value
                End Set
            End Property

            Private _incluisub As Boolean
            Public Property IncluiSub() As Boolean
                Get
                    Return _incluisub
                End Get
                Set(ByVal value As Boolean)
                    _incluisub = value
                End Set
            End Property

            Private _qtdarqs As Integer
            Public ReadOnly Property QtdArqs() As Integer
                Get
                    Return _qtdarqs
                End Get
            End Property

            Event NotificaStatus()

            Private _status As String
            Public Property Status() As String
                Get
                    Return _status
                End Get
                Set(ByVal value As String)
                    _status = value
                End Set
            End Property

            Private _inicio As Date = Nothing
            Public Property Inicio() As Date
                Get
                    Return _inicio
                End Get
                Set(ByVal value As Date)
                    _inicio = value
                End Set
            End Property

            Private _termino As Date = Nothing
            Public Property Termino() As Date
                Get
                    Return _termino
                End Get
                Set(ByVal value As Date)
                    _termino = value
                End Set
            End Property

            Private _logdetalhado As Boolean = False
            Public Property LogDetalhado() As Boolean
                Get
                    Return _logdetalhado
                End Get
                Set(ByVal value As Boolean)
                    _logdetalhado = value
                End Set
            End Property



            Private _log As New StringBuilder
            Public Property Log() As StringBuilder
                Get
                    Return _log
                End Get
                Set(ByVal value As StringBuilder)
                    _log = value
                End Set
            End Property

            Private Sub Trata(ByVal Arquivo As String, ByVal DirOrigem As String, ByVal DirDestino As String, Optional ByVal ArquivoDest As String = "")
                Try
                    If ArquivoDest = "" Then
                        ArquivoDest = Arquivo
                    End If

                    Dim ArqO As String = FileExpr(DirOrigem, Arquivo)
                    Dim ArqD As String = FileExpr(DirDestino, ArquivoDest)


                    Dim Tratou As Boolean = False
                    If _ListaApagar.Length > 0 Then
                        For Each Item As String In _ListaApagar
                            Try
                                If Item <> "" AndAlso Arquivo Like Item Then
                                    Try
                                        If System.IO.File.Exists(ArqO) Then
                                            System.IO.File.SetAttributes(ArqO, FileAttributes.Archive)
                                            System.IO.File.Delete(ArqO)
                                            RegLog("Eliminado arquivo " & ArqO & " por corresponder à máscara '" & Item & "'")
                                        End If
                                    Catch ex As Exception
                                        RegLog("[FALHA] " & ex.Message & " ao tentar excluir arquivo " & ArqO)
                                    End Try
                                    Try
                                        If System.IO.File.Exists(ArqD) Then
                                            System.IO.File.SetAttributes(ArqD, FileAttributes.Archive)
                                            System.IO.File.Delete(ArqD)
                                            RegLog("Eliminado arquivo " & ArqD & " por corresponder à máscara '" & Item & "'")
                                        End If
                                    Catch ex As Exception
                                        RegLog("[FALHA] " & ex.Message & " ao tentar excluir arquivo " & ArqD)
                                    End Try
                                    Tratou = True
                                    Exit For
                                End If
                            Catch ex As Exception
                                RegLog("[FALHA] " & ex.Message & " ao tentar validar expressão " & Item & " para arquivo " & Arquivo)
                            End Try
                        Next
                    End If

                    If Not Tratou Then
                        Dim AtribOrigem As New System.IO.FileInfo(ArqO)
                        Dim AtribDestino As New System.IO.FileInfo(ArqD)
                        If AtribOrigem.Exists Then
                            Dim TempoDif As Boolean = False
                            Try
                                TempoDif = AtribOrigem.LastWriteTime <> AtribDestino.LastWriteTime
                            Catch ex As Exception
                                RegLog("[FALHA] " & ex.Message & " ao tentar obter lastwritetime de " & ArqD)
                            End Try

                            If (Not AtribDestino.Exists) OrElse TempoDif OrElse (AtribOrigem.Length <> AtribDestino.Length) Then
                                Try
                                    FileCopy(ArqO, ArqD)
                                Catch EX As System.IO.DirectoryNotFoundException
                                    Try
                                        CriaDir(DirDestino)
                                        FileCopy(ArqO, ArqD)
                                    Catch Ex2 As Exception
                                        RegLog("[FALHA] " & Ex2.Message & " ao tentar copiar " & ArqO & " para " & ArqD)
                                    End Try
                                Catch ex As Exception
                                    RegLog("[FALHA] " & ex.Message & " ao tentar copiar " & ArqO & " para " & ArqD)
                                End Try
                            End If
                        End If
                    End If
                Catch ex As Exception
                    RegLog("[FALHA] " & ex.Message & " ao tentar copiar " & Arquivo & " para " & ArquivoDest)
                End Try
            End Sub

            Private Sub RegLog(ByVal Texto As String)
                Log.AppendLine(Format(Now, "ddd HH:mm:ss") & " - " & Texto)
            End Sub


            Private Sub Apaga(ByVal Arquivo As String, ByVal Diretorio As String)
                Try
                    Dim Arq As String = FileExpr(Diretorio, Arquivo)
                    Kill(Arq)
                    If LogDetalhado Then
                        RegLog("Apagou " & Arq)
                    End If
                Catch Ex As Exception
                    RegLog("[FALHA] " & Ex.Message & " ao tentar excluir " & Arquivo & " do diretório " & Diretorio)
                End Try
            End Sub

            Private Sub CriaDir(ByVal Diretorio As String)
                Try
                    MkDir(Diretorio)
                    If LogDetalhado Then
                        RegLog("Criou " & Diretorio)
                    End If
                Catch ex As Exception
                    RegLog("[FALHA] " & ex.Message & " ao tentar criar diretório " & Diretorio)
                End Try
            End Sub

            Dim UltNotif As String = ""
            Private Sub Notifica(ByVal Texto As String, Optional ByVal Forcar As Boolean = False)
                Dim Notif As String = Format(Now, "ss")
                If UltNotif <> Notif Or Forcar Then
                    Status = Texto
                    UltNotif = Notif
                    RaiseEvent NotificaStatus()
                End If
            End Sub

            Public Sub Executa()
                Try
                    Inicio = Now
                    Dim Ocorr As String = "Início de replicação de " & DirOrigem & " para " & DirDestino
                    Notifica(Ocorr, True)
                    RegLog(Ocorr)

                    Executa(DirOrigem, DirDestino)
                    Termino = Now

                    Ocorr = "Término (" & QtdArqs & Pl(QtdArqs, " arquivo") & " | duração: " & ExibeSegs(DateDiff(DateInterval.Second, Inicio, Termino), ExibeSegsOpc.xh_ymin_zseg) & ")"
                    Notifica(Ocorr, True)
                    RegLog(Ocorr)
                Catch ex As Exception
                    RegLog("[FALHA] " & ex.Message & " ao tentar executar sincronização entre " & DirOrigem & " e " & DirDestino)
                End Try
            End Sub

            Private DirBloqueado() As String = {"$RECYCLE.BIN", "System Volume Information"}
            Private Function Bloqueado(ByVal Caminho As String)
                Try
                    Dim Disco As String = System.IO.Path.GetPathRoot(Caminho)
                    Caminho = Mid(Caminho, Len(Disco) + 1) & "\"
                    For Each Item As String In DirBloqueado
                        If Caminho.StartsWith(Item & "\") Then
                            Return True
                        End If
                    Next
                Catch ex As Exception
                    RegLog("[FALHA] " & ex.Message & " ao verificar se caminho " & Caminho & " está na lista de itens bloqueados")
                End Try
                Return False
            End Function

            Private Sub Garante(ByVal Arquivo As String, ByVal DirOrigem As String, ByVal DirDestino As String, Optional ByVal ArquivoDest As String = "")
                Try
                    If ArquivoDest = "" Then
                        ArquivoDest = Arquivo
                    End If

                    If Not System.IO.File.Exists(FileExpr(DirOrigem, Arquivo)) AndAlso System.IO.File.Exists(FileExpr(DirDestino, Arquivo)) Then
                        Apaga(ArquivoDest, DirDestino)
                    End If
                Catch ex As Exception
                    RegLog("[FALHA] " & ex.Message & " ao buscar garantias de igualdade entre origem " & DirOrigem & "..." & Arquivo & " e " & DirDestino & "..." & ArquivoDest)
                End Try
            End Sub

            Private _ListaApagar() As String = {}

            Private Sub Executa(ByVal Origem As String, ByVal Destino As String)
                Try
                    _ListaApagar = Split(_ApagarQuandoEncontrar, vbCrLf)
                    If _ListaApagar.Length = 1 AndAlso Trim(_ListaApagar(0)) = "" Then
                        _ListaApagar = New String() {}
                    End If

                    ' garante todos os arquivos da origem no destino
                    If System.IO.Directory.Exists(Origem) Then


                        For Each Arq As String In System.IO.Directory.GetFiles(Origem)
                            If Not Bloqueado(Arq) Then
                                Notifica(Origem)

                                Dim ArqA As String = System.IO.Path.GetFileName(Arq)
                                Trata(ArqA, Origem, Destino)
                                _qtdarqs += 1
                            End If
                        Next

                        ' garante que não tenha nenhum a mais
                        If System.IO.Directory.Exists(Destino) Then
                            For Each Arq As String In System.IO.Directory.GetFiles(Destino)
                                If Not Bloqueado(Arq) Then
                                    Notifica(Arq)

                                    Dim ArqA As String = System.IO.Path.GetFileName(Arq)
                                    Garante(ArqA, Origem, Destino)
                                End If
                            Next
                        Else
                            CriaDir(Destino)
                        End If

                        ' diretório que existem na origem
                        For Each Dir As String In System.IO.Directory.GetDirectories(Origem)
                            If Not Bloqueado(Dir) Then
                                Notifica(Dir)

                                Dim DirA As String = System.IO.Path.GetFileName(Dir)
                                If IncluiSub Then
                                    Executa(Dir, FileExpr(Destino, DirA))
                                End If
                            End If
                        Next

                        ' diretórios existentes no destino sem origem
                        For Each Dir As String In System.IO.Directory.GetDirectories(Destino)
                            If Not Bloqueado(Dir) Then
                                Notifica(Dir)

                                Dim DirA As String = System.IO.Path.GetFileName(Dir)
                                If Not System.IO.Directory.Exists(FileExpr(Origem, DirA)) Then
                                    ApagaDir(Dir)
                                End If
                            End If
                        Next
                    Else
                        Dim OrigArq As String = System.IO.Path.GetFileName(Origem)
                        If OrigArq <> "" Then
                            Dim OrigSemArq As String = System.IO.Path.GetDirectoryName(Origem)
                            Dim DestArq As String = System.IO.Path.GetFileName(Destino)
                            Dim DestSemArq As String = System.IO.Path.GetDirectoryName(Destino)

                            Trata(OrigArq, OrigSemArq, DestSemArq, DestArq)
                            _qtdarqs += 1
                            Garante(OrigArq, OrigSemArq, DestSemArq, DestArq)

                        End If
                    End If
                Catch ex As Exception
                    RegLog("[FALHA] " & ex.Message & " ao executar sincronização entre origem " & Origem & " e destino " & Destino)
                End Try
            End Sub

            Sub ApagaDir(ByVal Diretorio As String)
                Try
                    System.IO.Directory.Delete(Diretorio, True)
                    If LogDetalhado Then
                        RegLog("Apagou " & Diretorio)
                    End If
                Catch EX As Exception
                    RegLog("[FALHA] ao apagar diretório " & Diretorio & ": " & EX.Message)
                End Try
            End Sub

        End Class

        Public Shared Function Pl(ByVal Numero As Object, ByVal Singular As String, Optional ByVal Plural As String = "") As String
            Return IIf(Numero = 1, Singular, NZV(Plural, Singular & IIf(Char.IsLower(Microsoft.VisualBasic.Right(Singular, 1)), "s", "S")))
        End Function


    End Class

End Namespace