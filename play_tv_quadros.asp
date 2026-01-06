<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 18/12/2025               -->
<!-- CODIGO_ARQUIVO: LSCWRLCDJG          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Dashboard Vendas - Modo TV Completo</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        :root {
            --color-primary: #4361ee;
            --color-secondary: #3a0ca3;
            --color-success: #4cc9f0;
            --color-warning: #f72585;
            --color-excelente: #28a745;
            --color-bom: #17a2b8;
            --color-medio: #ffc107;
            --color-baixo: #fd7e14;
            --color-dark: #0a0a0a;
            --color-light: #ffffff;
            --card-height: 45vh;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            background: linear-gradient(135deg, #000000 0%, #1a1a2e 100%);
            color: var(--color-light);
            font-family: 'Segoe UI', 'Roboto', sans-serif;
            overflow: hidden;
            height: 100vh;
            width: 100vw;
            position: relative;
        }

        /* Background */
        .bg-gradient {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: 
                radial-gradient(circle at 20% 50%, rgba(67, 97, 238, 0.15) 0%, transparent 50%),
                radial-gradient(circle at 80% 20%, rgba(76, 201, 240, 0.1) 0%, transparent 50%),
                radial-gradient(circle at 40% 80%, rgba(247, 37, 133, 0.1) 0%, transparent 50%);
            z-index: -1;
        }

        /* Container principal */
        .dashboard-container {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(420px, 1fr));
            grid-auto-rows: var(--card-height);
            gap: 20px;
            padding: 20px;
            align-content: center;
            justify-content: center;
            z-index: 1;
            transition: all 0.8s ease;
        }

        /* Modo normal */
        .normal-mode .card {
            opacity: 1;
            transform: scale(1);
            transition: all 0.6s ease;
        }

        /* Modo parada - APENAS 1 CARD */
        .stop-mode {
            grid-template-columns: 1fr !important;
            grid-template-rows: 1fr !important;
            align-items: center;
            justify-items: center;
        }

        .stop-mode .card {
            opacity: 0;
            transform: scale(0.5);
            pointer-events: none;
            transition: all 0.6s ease;
        }

        .stop-mode .card.active {
            opacity: 1;
            transform: scale(1.2);
            pointer-events: auto;
            width: 90%;
            height: 85vh;
            animation: focusAppear 0.8s cubic-bezier(0.34, 1.56, 0.64, 1) forwards;
        }

        @keyframes focusAppear {
            0% { 
                opacity: 0;
                transform: scale(0.3) rotateY(90deg);
                filter: blur(20px);
            }
            100% { 
                opacity: 1;
                transform: scale(1.2) rotateY(0deg);
                filter: blur(0);
            }
        }

        /* Cards */
        .card {
            background: linear-gradient(145deg, rgba(30, 30, 46, 0.95), rgba(20, 20, 36, 0.98));
            border-radius: 25px;
            border: 3px solid;
            position: relative;
            overflow: hidden;
            padding: 25px;
            box-shadow: 
                0 15px 50px rgba(0, 0, 0, 0.5),
                inset 0 1px 0 rgba(255, 255, 255, 0.1);
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            cursor: pointer;
            min-height: var(--card-height);
            backdrop-filter: blur(15px);
        }

        /* Bordas coloridas */
        .card.vgv { border-color: var(--color-primary); }
        .card.meta { border-color: var(--color-secondary); }
        .card.ticket { border-color: var(--color-success); }
        .card.unidades { border-color: var(--color-warning); }
        .card.top-ano { border-color: #7209b7; }
        .card.top-mes { border-color: #4cc9f0; }
        .card.top-trimestre { border-color: #f72585; }
        .card.top-semestre { border-color: #ff9e00; }

        /* Conteúdo do card */
        .card-content {
            width: 100%;
            height: 100%;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            align-items: center;
        }

        .card-header {
            width: 100%;
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 15px;
        }

        .card-icon {
            font-size: 3rem;
            opacity: 0.9;
            filter: drop-shadow(0 0 15px currentColor);
        }

        .period-badge {
            padding: 6px 15px;
            border-radius: 20px;
            font-size: 0.8rem;
            font-weight: 700;
            background: rgba(255, 255, 255, 0.15);
            border: 2px solid rgba(255, 255, 255, 0.3);
            letter-spacing: 1px;
        }

        .period-ano { background: rgba(114, 9, 183, 0.3); border-color: #7209b7; }
        .period-mes { background: rgba(76, 201, 240, 0.3); border-color: #4cc9f0; }
        .period-trimestre { background: rgba(247, 37, 133, 0.3); border-color: #f72585; }
        .period-semestre { background: rgba(255, 158, 0, 0.3); border-color: #ff9e00; }

        .card-title {
            font-size: 1.8rem;
            font-weight: 800;
            margin: 15px 0;
            color: rgba(255, 255, 255, 0.95);
            text-transform: uppercase;
            letter-spacing: 1.5px;
            line-height: 1.2;
        }

        .card-value {
            font-size: 3.8rem;
            font-weight: 900;
            margin: 20px 0;
            text-shadow: 
                0 0 30px currentColor,
                0 0 60px rgba(255, 255, 255, 0.3);
            line-height: 1;
            background: linear-gradient(45deg, var(--color-light), var(--color-success));
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
        }

        .card-details {
            width: 100%;
            margin-top: 20px;
            padding: 15px;
            background: rgba(255, 255, 255, 0.05);
            border-radius: 15px;
        }

        .detail-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 0;
            border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        }

        .detail-row:last-child {
            border-bottom: none;
        }

        .detail-label {
            font-size: 1rem;
            color: rgba(255, 255, 255, 0.7);
        }

        .detail-value {
            font-size: 1.2rem;
            font-weight: 700;
            color: var(--color-success);
        }

        .card-footer {
            width: 100%;
            margin-top: 20px;
        }

        .progress-container {
            width: 100%;
            height: 12px;
            background: rgba(255, 255, 255, 0.15);
            border-radius: 6px;
            margin: 15px 0;
            overflow: hidden;
            position: relative;
        }

        .progress-bar {
            height: 100%;
            border-radius: 6px;
            position: relative;
            transition: width 1.5s cubic-bezier(0.34, 1.56, 0.64, 1);
            background: linear-gradient(90deg, 
                var(--color-primary),
                var(--color-success));
        }

        .card-subtitle {
            font-size: 1rem;
            color: rgba(255, 255, 255, 0.7);
            margin-top: 10px;
            line-height: 1.4;
        }

        /* Top performers */
        .top-performer {
            width: 100%;
            padding: 20px;
            margin: 10px 0;
            background: linear-gradient(90deg, 
                rgba(255, 255, 255, 0.08),
                rgba(255, 255, 255, 0.03));
            border-radius: 15px;
            border-left: 6px solid;
            animation: slideIn 0.5s ease-out;
        }

        .top-performer.gold { border-left-color: #FFD700; }
        .top-performer.silver { border-left-color: #C0C0C0; }
        .top-performer.bronze { border-left-color: #CD7F32; }

        .performer-rank {
            font-size: 2.5rem;
            font-weight: 900;
            margin-bottom: 10px;
            text-shadow: 0 0 15px currentColor;
        }

        .performer-rank.gold { color: #FFD700; }
        .performer-rank.silver { color: #C0C0C0; }
        .performer-rank.bronze { color: #CD7F32; }

        .performer-name {
            font-size: 1.5rem;
            font-weight: 700;
            margin-bottom: 8px;
            color: rgba(255, 255, 255, 0.95);
        }

        .performer-value {
            font-size: 2rem;
            font-weight: 800;
            color: var(--color-success);
            margin: 10px 0;
        }

        .performer-change {
            font-size: 1.1rem;
            font-weight: 700;
            color: #4cd964;
        }

        /* Controles */
        .controls {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 1000;
            display: flex;
            gap: 20px;
            align-items: center;
        }

        .counter {
            background: rgba(10, 10, 20, 0.9);
            color: white;
            padding: 15px 25px;
            border-radius: 25px;
            font-size: 1.2rem;
            font-weight: bold;
            border: 3px solid var(--color-primary);
            backdrop-filter: blur(15px);
            min-width: 180px;
            text-align: center;
            box-shadow: 0 15px 40px rgba(0, 0, 0, 0.4);
        }

        .counter .number {
            font-size: 1.8rem;
            color: var(--color-success);
            margin-left: 10px;
            text-shadow: 0 0 20px var(--color-success);
        }

        /* Título principal */
        .main-title {
            position: fixed;
            top: 20px;
            left: 20px;
            font-size: 2.5rem;
            font-weight: 900;
            background: linear-gradient(45deg, 
                var(--color-primary),
                var(--color-success),
                var(--color-warning));
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
            z-index: 1000;
            text-shadow: 0 0 50px rgba(67, 97, 238, 0.3);
        }

        /* Timer */
        .timer {
            position: fixed;
            bottom: 30px;
            left: 50%;
            transform: translateX(-50%);
            font-size: 1.2rem;
            color: rgba(255, 255, 255, 0.7);
            background: rgba(0, 0, 0, 0.3);
            padding: 10px 25px;
            border-radius: 20px;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }

        /* Animações */
        @keyframes slideIn {
            from { 
                opacity: 0;
                transform: translateX(-30px);
            }
            to { 
                opacity: 1;
                transform: translateX(0);
            }
        }

        .card-exit {
            animation: cardExit 0.8s forwards cubic-bezier(0.34, 1.56, 0.64, 1);
        }

        @keyframes cardExit {
            0% { 
                opacity: 1;
                transform: scale(1) rotate(0deg);
            }
            100% { 
                opacity: 0;
                transform: scale(0) rotate(360deg);
            }
        }

        .card-enter {
            animation: cardEnter 0.8s forwards cubic-bezier(0.34, 1.56, 0.64, 1);
        }

        @keyframes cardEnter {
            0% { 
                opacity: 0;
                transform: scale(0) rotate(-360deg);
            }
            100% { 
                opacity: 1;
                transform: scale(1) rotate(0deg);
            }
        }
    </style>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
</head>
<body>
    <div class="bg-gradient"></div>
    
    <div class="main-title">
        <i class="fas fa-chart-line"></i> DASHBOARD VENDAS
    </div>

    <div class="controls">
        <div class="counter">
            <i class="fas fa-sync-alt"></i> CICLO: 
            <span class="number" id="cycleCount">0</span>/5
        </div>
    </div>

    <div class="timer" id="timerDisplay">Próximo: <span id="timeLeft">3s</span></div>

    <div class="dashboard-container normal-mode" id="dashboardContainer">
        <!-- Cards serão gerados dinamicamente -->
    </div>

    <script>
        // Configurações
        const CONFIG = {
            moveInterval: 30000,          // 3 segundos entre movimentos
            movesBeforeStop: 5,          // 5 movimentos antes da parada
            stopDuration: 5000,          // 5 segundos em modo parada (APENAS 1 CARD)
            totalCards: 8
        };

        // Dados com período definido
        const PERFORMANCE_DATA = {
            vgvAnual: { 
                value: 24568432.56, 
                label: "VGV ANUAL", 
                icon: "fas fa-calendar-alt", 
                type: "vgv",
                period: "ano",
                change: "+18.45%",
                details: [
                    { label: "Meta Anual", value: "R$ 22.5M" },
                    { label: "Atingimento", value: "109.2%" },
                    { label: "Crescimento", value: "+18.45%" }
                ]
            },
            metaMensal: { 
                value: 94.25, 
                label: "META MENSAL", 
                icon: "fas fa-bullseye", 
                type: "meta", 
                suffix: "%",
                period: "mês",
                change: "+12.75%",
                details: [
                    { label: "Meta Mensal", value: "R$ 1.8M" },
                    { label: "Realizado", value: "R$ 1.96M" },
                    { label: "Excedente", value: "R$ 160K" }
                ]
            },
            ticketMedio: { 
                value: 325478.92, 
                label: "TICKET MÉDIO", 
                icon: "fas fa-crown", 
                type: "ticket",
                period: "trimestre",
                change: "+22.83%",
                details: [
                    { label: "Período", value: "Último Trimestre" },
                    { label: "Variação", value: "+22.83%" },
                    { label: "Comparativo", value: "R$ 265K" }
                ]
            },
            unidades: { 
                value: 78, 
                label: "UNIDADES VENDIDAS", 
                icon: "fas fa-boxes", 
                type: "unidades",
                period: "mês",
                change: "+15.62%",
                details: [
                    { label: "Meta Mensal", value: "65 unidades" },
                    { label: "Atingimento", value: "120.0%" },
                    { label: "Crescimento", value: "+15.62%" }
                ]
            }
        };

        // Top performers por período
        const TOP_PERFORMERS = {
            ano: {
                label: "TOP DO ANO",
                icon: "fas fa-trophy",
                type: "top-ano",
                period: "ano",
                performers: [
                    { 
                        rank: 1, 
                        name: "DIRETORIA CENTRO", 
                        value: 2456843.21, 
                        change: "+24.56%",
                        medal: "gold"
                    }
                ]
            },
            mes: {
                label: "TOP DO MÊS",
                icon: "fas fa-star",
                type: "top-mes",
                period: "mês",
                performers: [
                    { 
                        rank: 1, 
                        name: "CARLOS SILVA", 
                        value: 1254789.32, 
                        change: "+28.42%",
                        medal: "gold"
                    }
                ]
            },
            trimestre: {
                label: "TOP DO TRIMESTRE",
                icon: "fas fa-medal",
                type: "top-trimestre",
                period: "trimestre",
                performers: [
                    { 
                        rank: 1, 
                        name: "RESIDENCIAL SOLAR", 
                        value: 4856321.78, 
                        change: "+40.15%",
                        medal: "gold"
                    }
                ]
            },
            semestre: {
                label: "TOP DO SEMESTRE",
                icon: "fas fa-award",
                type: "top-semestre",
                period: "semestre",
                performers: [
                    { 
                        rank: 1, 
                        name: "GERÊNCIA ALPHA", 
                        value: 1854321.45, 
                        change: "+35.72%",
                        medal: "gold"
                    }
                ]
            }
        };

        // Estado da aplicação
        let state = {
            cycleCount: 0,
            isStopMode: false,
            currentStopCard: null,
            intervalId: null,
            stopTimeoutId: null,
            cards: []
        };

        // Elementos DOM
        const dashboardContainer = document.getElementById('dashboardContainer');
        const cycleCountElement = document.getElementById('cycleCount');
        const timerDisplay = document.getElementById('timerDisplay');
        const timeLeftElement = document.getElementById('timeLeft');

        // Timer
        let timeLeft = CONFIG.moveInterval / 1000;
        let timerInterval;

        // Inicialização
        function init() {
            createCards();
            renderCards();
            startMovementCycle();
            startTimer();
        }

        // Cria cards com períodos definidos
        function createCards() {
            // Cards de métricas
            Object.values(PERFORMANCE_DATA).forEach((data, index) => {
                state.cards.push({
                    id: `metric-${index}`,
                    data: data,
                    element: null,
                    isTopPerformer: false
                });
            });
            
            // Cards de top performers (apenas 1 por período)
            Object.values(TOP_PERFORMERS).forEach((data, index) => {
                state.cards.push({
                    id: `top-${index}`,
                    data: data,
                    element: null,
                    isTopPerformer: true
                });
            });
        }

        // Renderiza todos os cards
        function renderCards() {
            dashboardContainer.innerHTML = '';
            
            state.cards.forEach((card, index) => {
                const cardElement = document.createElement('div');
                cardElement.className = `card ${card.data.type}`;
                cardElement.style.animationDelay = `${index * 0.1}s`;
                
                if (state.isStopMode) {
                    // No modo parada, só mostra se for o card ativo
                    if (card.id === state.currentStopCard) {
                        cardElement.classList.add('active');
                    } else {
                        cardElement.style.display = 'none';
                    }
                } else {
                    cardElement.classList.add('card-enter');
                }
                
                if (card.isTopPerformer) {
                    cardElement.innerHTML = createTopPerformerCard(card);
                } else {
                    cardElement.innerHTML = createMetricCard(card);
                }
                
                dashboardContainer.appendChild(cardElement);
                card.element = cardElement;
            });
        }

        // Cria card de métrica com período
        function createMetricCard(card) {
            const value = formatValue(card.data.value, card.data.suffix);
            const progressWidth = card.data.suffix === '%' ? card.data.value : Math.min(100, (card.data.value / 30000000) * 100);
            const periodClass = `period-${card.data.period}`;
            const periodLabel = getPeriodLabel(card.data.period);
            
            return `
                <div class="card-content">
                    <div class="card-header">
                        <i class="${card.data.icon} card-icon"></i>
                        <span class="period-badge ${periodClass}">${periodLabel}</span>
                    </div>
                    
                    <div class="card-title">${card.data.label}</div>
                    
                    <div class="card-value">${value}</div>
                    
                    <div class="card-details">
                        ${card.data.details.map(detail => `
                            <div class="detail-row">
                                <span class="detail-label">${detail.label}:</span>
                                <span class="detail-value">${detail.value}</span>
                            </div>
                        `).join('')}
                    </div>
                    
                    <div class="card-footer">
                        <div class="progress-container">
                            <div class="progress-bar" style="width: ${progressWidth}%"></div>
                        </div>
                        <div class="card-subtitle">
                            Crescimento: <span style="color: #4cd964">${card.data.change}</span>
                        </div>
                    </div>
                </div>
            `;
        }

        // Cria card de top performer (apenas 1)
        function createTopPerformerCard(card) {
            const performer = card.data.performers[0];
            const periodClass = `period-${card.data.period}`;
            const periodLabel = getPeriodLabel(card.data.period);
            
            return `
                <div class="card-content">
                    <div class="card-header">
                        <i class="${card.data.icon} card-icon"></i>
                        <span class="period-badge ${periodClass}">${periodLabel}</span>
                    </div>
                    
                    <div class="card-title">${card.data.label}</div>
                    
                    <div class="top-performer ${performer.medal}">
                        <div class="performer-rank ${performer.medal}">#${performer.rank}</div>
                        <div class="performer-name">${performer.name}</div>
                        <div class="performer-value">${formatCurrency(performer.value)}</div>
                        <div class="performer-change">${performer.change}</div>
                    </div>
                    
                    <div class="card-footer">
                        <div class="card-subtitle">
                            Melhor desempenho do ${periodLabel.toLowerCase()}
                        </div>
                    </div>
                </div>
            `;
        }

        // Funções auxiliares
        function formatValue(value, suffix = '') {
            if (suffix === '%') {
                return `${value.toFixed(2)}%`;
            } else if (value >= 1000000) {
                return `R$ ${(value / 1000000).toFixed(2).replace('.', ',')}M`;
            } else if (value >= 1000) {
                return `R$ ${(value / 1000).toFixed(2).replace('.', ',')}K`;
            }
            return `R$ ${value.toFixed(2).replace('.', ',')}`;
        }

        function formatCurrency(value) {
            if (value >= 1000000) {
                return `R$ ${(value / 1000000).toFixed(2).replace('.', ',')}M`;
            } else if (value >= 1000) {
                return `R$ ${(value / 1000).toFixed(2).replace('.', ',')}K`;
            }
            return `R$ ${value.toFixed(2).replace('.', ',')}`;
        }

        function getPeriodLabel(period) {
            const labels = {
                'ano': 'ANO',
                'mes': 'MÊS',
                'trimestre': 'TRIMESTRE',
                'semestre': 'SEMESTRE'
            };
            return labels[period] || period.toUpperCase();
        }

        function getColorForChange(change) {
            return change.startsWith('+') ? '#4cd964' : '#ff6b6b';
        }

        // Inicia o ciclo de movimento
        function startMovementCycle() {
            state.intervalId = setInterval(moveCards, CONFIG.moveInterval);
        }

        // Move e embaralha os cards
        function moveCards() {
            if (state.isStopMode) return;
            
            state.cycleCount++;
            cycleCountElement.textContent = state.cycleCount;
            
            // Anima saída dos cards
            state.cards.forEach(card => {
                if (card.element) {
                    card.element.classList.add('card-exit');
                }
            });
            
            // Embaralha e re-renderiza
            setTimeout(() => {
                shuffleCards();
                renderCards();
                dashboardContainer.classList.remove('stop-mode');
                dashboardContainer.classList.add('normal-mode');
                
                // Verifica se deve entrar em modo parada
                if (state.cycleCount % CONFIG.movesBeforeStop === 0 && !state.isStopMode) {
                    enterStopMode();
                }
                
            }, 800);
        }

        // Embaralha os cards
        function shuffleCards() {
            for (let i = state.cards.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [state.cards[i], state.cards[j]] = [state.cards[j], state.cards[i]];
            }
        }

        // Entra no modo parada (APENAS 1 CARD)
        function enterStopMode() {
            state.isStopMode = true;
            clearInterval(state.intervalId);
            
            // Seleciona aleatoriamente 1 card
            const availableCards = state.cards.filter(card => card.element);
            if (availableCards.length > 0) {
                const randomIndex = Math.floor(Math.random() * availableCards.length);
                state.currentStopCard = availableCards[randomIndex].id;
            }
            
            // Ativa modo parada
            setTimeout(() => {
                dashboardContainer.classList.remove('normal-mode');
                dashboardContainer.classList.add('stop-mode');
                renderCards();
                
                // Atualiza timer
                timerDisplay.style.display = 'block';
                timeLeft = CONFIG.stopDuration / 1000;
                updateTimerDisplay();
                
            }, 500);
            
            // Configura para sair do modo parada após 5 segundos
            state.stopTimeoutId = setTimeout(exitStopMode, CONFIG.stopDuration);
        }

        // Sai do modo parada
        function exitStopMode() {
            state.isStopMode = false;
            state.currentStopCard = null;
            
            // Remove modo parada
            dashboardContainer.classList.remove('stop-mode');
            dashboardContainer.classList.add('normal-mode');
            
            // Re-renderiza todos os cards
            renderCards();
            
            // Esconde timer
            timerDisplay.style.display = 'none';
            
            // Reinicia o ciclo após um pequeno delay
            setTimeout(() => {
                state.intervalId = setInterval(moveCards, CONFIG.moveInterval);
                timeLeft = CONFIG.moveInterval / 1000;
            }, 1000);
        }

        // Timer visual
        function startTimer() {
            timerInterval = setInterval(() => {
                if (state.isStopMode) {
                    timeLeft = Math.max(0, timeLeft - 1);
                    updateTimerDisplay();
                    
                    if (timeLeft === 0) {
                        timeLeft = CONFIG.moveInterval / 1000;
                    }
                } else {
                    timeLeft = timeLeft > 0 ? timeLeft - 1 : CONFIG.moveInterval / 1000;
                    updateTimerDisplay();
                }
            }, 1000);
        }

        function updateTimerDisplay() {
            timeLeftElement.textContent = `${timeLeft}s`;
        }

        // Atualiza dados periodicamente
        setInterval(() => {
            if (!state.isStopMode) {
                // Simula pequenas variações nos dados
                Object.values(PERFORMANCE_DATA).forEach(data => {
                    if (data.suffix === '%') {
                        data.value = Math.min(100, 
                            data.value + (Math.random() * 0.5 - 0.25)
                        );
                        data.value = parseFloat(data.value.toFixed(2));
                    } else {
                        const variation = data.value * 0.02;
                        data.value += (Math.random() * variation * 2 - variation);
                        data.value = parseFloat(data.value.toFixed(2));
                    }
                    
                    const changeValue = Math.random() * 0.5 + 0.1;
                    data.change = `+${changeValue.toFixed(2)}%`;
                });
                
                renderCards();
            }
        }, 10000);

        // Inicializa
        document.addEventListener('DOMContentLoaded', init);
    </script>
</body>
</html>