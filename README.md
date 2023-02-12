# Google-Admin-Create-User-Api
Criando Usuários pela Api do google admin com planilhas google
# Por que?
Devido a necessidade de criação automatizada de emails no google, foi criado esta proposta.
# Como foi feito?
Usando Api do Google Admin Directory com planilhas Google
# Como funciona?
Em uma planilha adiciona-se o nome completo em uma coluna, após isso a função roda automatico, pega o valor nome quebra em pedaços e junta primeiro e ultimo nome, concatena com email e confere se já existe em uma coluna da planilha, se sim, pega o penultimo nome, e assim por diante.
# O que falta?
Os nomes devem vir de outro sistema, por Api ou banco de dados, e preencher sempre o lastrow da coluna nome. Assim sempre que for preenchido roda-se a função.
# Adendo
Devemos nos atentar que se já existem emails cadastrador devemos adicionar na planilha Email e Nome.
![Screenshot_3](https://user-images.githubusercontent.com/97992826/218317004-dd06140e-b9bd-4c7b-a2a3-19d68031937b.png)
