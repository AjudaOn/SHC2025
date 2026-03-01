// document.addEventListener('DOMContentLoaded', function() {
//     // Adiciona o ouvinte de eventos ao link "Nova reserva"
//     var carregarReservasLink = document.getElementById('carregar_reservas');
//     if (carregarReservasLink) {
//         carregarReservasLink.addEventListener('click', function(event) {
//             event.preventDefault(); // Previne o comportamento padrão do link
//             carregarReservasEConfigurarEventos();
//         });
//     }
// });
// function inicializarReservas() {
//     // Seu código para inicializar reservas vai aqui
// }

// document.addEventListener('DOMContentLoaded', function() {
//     var carregarReservasLink = document.getElementById('carregar_reservas');
//     if (carregarReservasLink) {
//         carregarReservasLink.addEventListener('click', function(event) {
//             event.preventDefault();
//             carregarReservasEConfigurarEventos();
//         });
//     }
// });

// function carregarReservasEConfigurarEventos() {
//     // Substitua '/caminho/para/reservas/' pelo caminho correto no seu app Django
//     fetch('/reservas/reservas/') 
//         .then(response => response.text())
//         .then(html => {
//             document.getElementById('conteudoDinamico').innerHTML = html;
//             configurarManipuladorQtdeHosp();
//             // configurarManipuladorQtdeQuartos(); // Chame esta função se necessário
//         });
// }

// function configurarManipuladorQtdeHosp() {
//     var qtdeHospInput = document.getElementById('id_qtde_hosp');
//     if (qtdeHospInput) {
//         qtdeHospInput.addEventListener('change', toggleAcompanhanteDivs);
//         toggleAcompanhanteDivs(); // Aplica a lógica inicialmente
//     }
// }

// function configurarManipuladorQtdeHosp() {
//     var qtdeHospInput = document.getElementById('id_qtde_hosp');
//     if (qtdeHospInput) {
//         // Trocando 'change' por 'input' para reação imediata a cada alteração
//         qtdeHospInput.addEventListener('input', function() {
//             // Sua lógica de tratamento aqui
//             toggleAcompanhanteDivs();
//         });
//     }
// }

// Nova função para configurar o ouvinte de eventos no campo qtde_quartos
// function configurarManipuladorQtdeQuartos() {
//     var qtdeQuartosInput = document.getElementById('inputQtde_quartos');
//     if (qtdeQuartosInput) {
//         qtdeQuartosInput.addEventListener('focus', toggleAcompanhanteDivs); // Aciona quando o campo ganha foco
//         // Considerar a adição de "blur" se desejar reavaliar quando o usuário sai do campo
//         // qtdeQuartosInput.addEventListener('blur', toggleAcompanhanteDivs);
//     }
// }

function toggleAcompanhanteDivs() {
    var qtdeHospInput = document.getElementById('id_qtde_hosp');
    var qtdeHosp = parseInt(qtdeHospInput.value, 10) -1;
    console.log('Quantidade de Hóspedes:', qtdeHosp);
    
    // Loop para mostrar/esconder divs de acompanhantes baseado em qtde_hosp
    for (var i = 1; i <= 5; i++) {
        var div = document.getElementById('acompanhante' + i);
        if (div) {
            div.className = i <= qtdeHosp ? 'visible' : 'hidden';
        }
    }
    var qtdeHosp = parseInt(qtdeHospInput.value, 10);
    print(qtdeHosp);
    // Ajustando visibilidade para singular ou plural
    if (qtdeHosp === 1) {
        document.getElementById('singular').className = 'hidden';
        document.getElementById('plural').className = 'hidden';
        document.getElementById('linha').className = 'hidden';
    } else if (qtdeHosp === 2) {
        document.getElementById('singular').className = 'visible';
        document.getElementById('plural').className = 'hidden';
        document.getElementById('linha').className = 'visible';
    } else {
        document.getElementById('singular').className = 'hidden';
        document.getElementById('plural').className = 'visible';
        document.getElementById('linha').className = 'visible';
    }
}

// function carregarAguardandoAprovacaoEConfigurarEventos() {
//     // Substitua '/caminho/para/aguardando_aprovacao/' pelo caminho correto no seu app Django
//     fetch('/reservas/aguardando/') 
//         .then(response => response.text())
//         .then(html => {
//             document.getElementById('conteudoDinamico').innerHTML = html;
//             configurarManipuladorQtdeHosp();
//             configurarManipuladoresEventosFormulario();
//             // Chamada às funções que controlam a visibilidade das divs com base em qtde_hosp
//             toggleAcompanhanteDivs();
//             // Se houver mais lógicas específicas para este template, chame-as aqui
//         });
// }



// setTimeout(function() {
//     toggleAcompanhanteDivs();
// }, 100);







