var getCodigoAluno = function(ra) {
	return new Promise(function(resolve, reject) {
		var formData = new FormData();

		formData.append('anoLetivo', '0');
		formData.append('tipoConsultaFichaAluno', '1');
		formData.append('ra', String(ra));
		formData.append('digRa', '');
		formData.append('ufRa', 'SP');
		formData.append('nomeSocial', '');
		formData.append('dataNascimento', '');
		formData.append('nomeCompleto', '');
		formData.append('codigoDiretoria', '0');
		formData.append('codigoEscola', '0');
		formData.append('tipoEnsino', '0');
		formData.append('codigoTurma', '0');
		formData.append('numeroClasse', '');
		formData.append('nomeAluno', '');
		formData.append('nomeMaeFonetico', '');
		formData.append('nomeMae', '');
		formData.append('nomePai', '');
		formData.append('rg', '');
		formData.append('digRg', '');
		formData.append('ufRg', 'SP');
		formData.append('cpf', '');
		formData.append('NIS', '');
		formData.append('INEP', '');
		formData.append('codigoCertidao', '');
		formData.append('LivroNasc', '');
		formData.append('FolhaReg', '');
		formData.append('RegNasc', '');
		formData.append('CertidaoNova', 'true');
		formData.append('paginaAtual', '1');
		formData.append('ConsultaProgramas', 'False');
		formData.append('nomeFonetico', '');
		
		var xhr = new XMLHttpRequest();
		xhr.open('POST', 'https://sed.educacao.sp.gov.br/NCA/FichaAluno/ListaFichaAlunoParcial', true);
		xhr.send(formData);

		xhr.onreadystatechange = function () {
			if (xhr.readyState === 4) {
				if (xhr.status === 200) {
					resolve(JSON.stringify(xhr.responseText));
				} else {
					reject('Erro na requisição');
				}        
			}  
		}		
	});
}

var getEachRg = function(codigoAluno) {
	return new Promise(function(resolve, reject) {
		var formData = new FormData();

	    formData.append('codigoAluno', String(codigoAluno));
	    formData.append('editar', 'false');
		formData.append('ConsultaProgramas', 'False');
		
		var xhr = new XMLHttpRequest();
		xhr.open('POST', 'https://sed.educacao.sp.gov.br/NCA/FichaAluno/FichaAluno', true);
		xhr.send(formData);

		xhr.onreadystatechange = function () {
			if (xhr.readyState === 4) {
				if (xhr.status === 200) {
					resolve(JSON.stringify(xhr.responseText));
				} else {
					reject('Erro na requisição');
				}        
			}  
		} 
	});
}

function getRgs(alunos){
	var rgsFinal = alunos.map(function(aluno) {
		getCodigoAluno(aluno.ra)
		.then(function(response) {
			var n = response.search("DadosFichaAluno");
			var codigoAluno = response.slice(n + 16, n + 24);
						
			getEachRg(codigoAluno)
				.then(function(response) {     
					var nRg = response.search('form-control force-inline RG');
					var rg = response.slice(nRg + 136, nRg + 150);
					var nRgDig = response.search('form-control force-inline digRG');				
					var digRg = response.slice(nRgDig + 143, nRgDig + 145);
					var nState = response.search('sgUfRg');				
					var state = response.slice(nState + 96, nState + 98);
					if (typeof(digRg[1]) !== Number) {
						digRg = digRg[0];
					}
					var rgWithdigits = rg + '-' + digRg + '/' + state;
					console.log(aluno.index + '#' + rgWithdigits);
					return rgWithdigits;
				})
				.catch(function(error) {
					console.warn(error);
				})
		})
		.catch(function(error) {
			console.warn(error);
		})
	});
	return Promise.all(rgsFinal);
}

getRgs(
	
);
