Código criado por:
	Raphael do Espirito Santo Nascimento 	RA: 186305
		ME323 Turma B
	Thais Brito Del Picchia Faria		RA: 206043
		ME323 Turma c
	Maria Clara Falcão Cunha		RA:202755
		ME323 Turma B

IMPORTANTE:
	A subpasta 'fotos' deve estar na mesma pasta que a planilha excel
		(Exemplo: Se a planilha está em C:/Downloads, deve existir um caminho C:/Downloads/fotos)
	É necessário descompactar a pasta
	Caso somente a simulação automática for executada, não é necessário da pasta 'fotos'


COMO JOGAR:
	1. Escolha entre os três botões disponíveis, sendo eles:
		1.1. Começar um novo jogo:
			1.1.1. Clique em uma porta
			1.1.2. Decida se quer mudar de porta ou continuar na mesma
			1.1.3. Clique na porta que você decidir e torça para ganhar
		1.2. Começar um jogo automático com troca de portas:
			- 1000 simulações irão ocorrer, aperte esc para para-las a qualquer momento (pode causar erro)
		1.3. Começar um jogo automático sem troca de portas:
			- 1000 simulações irão ocorrer, aperte esc para para-las a qualquer momento (pode causar erro)
	2. Aperte o botão novamente para tentar denovo


SOBRE O CÓDIGO:
	1. O código foi escrito em linguagem VBA
		1.1. Para visualizá-lo, aperte Alt+F11 e clique duas vezes em Plan1(Plan1)
	2. Sobre o jogo manual:
		2.1. Ao clicar no botão 'Começar um novo jogo', as portas são fechadas e dois números aleatórios são gerados:
			2.1.1. O primeiro número varia de 1 a 3, sendo a porta que será escolhida para receber o prêmio
			2.1.2. O segundo número varia de 1 a 2, sendo utilizado para decidir a porta que será aberta caso o jogador escolha a porta premiada
		2.2. Quando o jogador clica em uma porta:
			2.2.1. Será aberta a porta sem o prêmio, se a porta escolhida pelo jogador também não for a premiada
			2.2.2. Será escolhida uma porta baseada no segundo número aleatório, para abrir uma das portas sem prêmio, caso o jogador tenha clicado na porta premiada
		2.3. Ao clicar novamente em uma porta
			2.3.1. A porta será aberta
		2.4. É possível abrir todas as portas, onde é averiguavel que existem 2 portas com uma cabra e 1 com um prêmio
	3. Sobre os jogos automáticos:
		3.1. Com troca de porta:
			3.1.1. Ao clicar no botão 'Começar um novo jogo automático com troca de portas'
				3.1.1.1. A tabela com os resultados da ultima simulação serão apagados
				3.1.1.2. A contagem do número da simulação é retornado para 0
				3.1.1.3. É exibido a mensagem sobre o número de simulações e como cancela-la
				3.1.1.4. Iniciado um loop de 1000 voltas:
					3.1.1.4.1. Será definido 3 números aleatórios
						3.1.1.4.1.1. O primeiro número varia de 1 a 3, sendo a porta que será escolhida para receber o prêmio
						3.1.1.4.1.2. O segundo número varia de 1 a 2, sendo utilizado para decidir a porta que será aberta caso o jogador escolha a porta premiada
						3.1.1.4.1.3. O terceiro número varia de 1 a 3, que será a porta escolhida pelo programa inicialmente
					3.1.1.4.2. Se a porta escolhida pelo terceiro número aleatório for a premiada, o programa irá utilizar do segundo número para escolher uma porta que será "aberta" e escolherá como porta final a outra restante
					3.1.1.4.3. Se a porta escolhida pelo terceiro número aleatório não for a premiada, o programa irá abrir a outra porta sem prêmio e então escolherá como porta final a porta restante, que está com o prêmio
					3.1.1.4.4. A célula da linha a ser editada é selecionada, para acompanhamento pelo usuário
					3.1.1.4.5. As colunas da linha correspondente ao número da simulação são atualizadas
					3.1.1.4.6. O loop é reiniciado
		3.2. Sem troca de porta:
			3.2.1. Ao clicar no botão 'Começar um novo jogo sem troca de portas'
				3.2.1.1. A tabela com os resultados da ultima simulação serão apagados
				3.2.1.2. A contagem do número da simulação é retornado para 0
				3.2.1.3. É exibido a mensagem sobre o número de simulações e como cancela-la
				3.2.1.4. Iniciado um loop de 1000 voltas:
					3.2.1.4.1. Será definido 2 números aleatórios
						3.2.1.4.1.1. O primeiro número varia de 1 a 3, sendo a porta que será escolhida para receber o prêmio
						3.2.1.4.1.2. O segundo número varia de 1 a 3, que será a porta escolhida pelo programa inicialmente, e como não há troca
							de portas, será o número da porta final
					3.2.1.4.2. A célula da linha a ser editada é selecionada, para acompanhamento pelo usuário
					3.2.1.4.3. As colunas da linha correspondente ao número da simulação são atualizadas
					3.2.1.4.4. O loop é reiniciado
