// ==========================================
// 1. CONFIGURAÇÃO DE TEMA (DARK/LIGHT)
// ==========================================
const setTheme = theme => {
    if (theme === 'auto' && window.matchMedia('(prefers-color-scheme: dark)').matches) {
        document.documentElement.setAttribute('data-bs-theme', 'dark');
    } else {
        document.documentElement.setAttribute('data-bs-theme', theme);
    }
    localStorage.setItem('theme', theme);

    const btnIcon = document.querySelector('.dropdown-toggle i');
    if (btnIcon) {
        if (theme === 'light') btnIcon.className = 'bi bi-sun-fill';
        else if (theme === 'dark') btnIcon.className = 'bi bi-moon-stars-fill';
        else btnIcon.className = 'bi bi-circle-half';
    }
};

const savedTheme = localStorage.getItem('theme') || 'auto';
setTheme(savedTheme);

document.querySelectorAll('[data-theme]').forEach(element => {
    element.addEventListener('click', (e) => {
        e.preventDefault();
        setTheme(element.getAttribute('data-theme'));
    });
});

// ==========================================
// 2. LÓGICA DO GERADOR DE CERTIFICADOS
// ==========================================
const canvas = document.getElementById('certCanvas');
const ctx = canvas.getContext('2d');
const configIds = ['yPos', 'maxWidth', 'fontSize', 'lineHeight'];
let baseImg = new Image();
let dados = [];

// Limpa "horas" duplicadas e garante que o texto seja tratado como string
const limparHoras = (v) => v ? v.toString().toLowerCase().replace(/horas/gi, '').trim() : "0";

// Template do Texto com tratamento de Caixa Alta para o Nome
const getTexto = (n, c, f, h) => `Certificamos que ${n.toString().toUpperCase()}, portadora do CPF ${c}, participou da ${f}, com carga horária total de ${limparHoras(h)} horas.`;

// Função para remover acentos apenas dos nomes de arquivos (evita erro no Windows)
const removerAcentos = (str) => {
    return str.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9_]/g, "_");
};

// Carregar Imagem Base
document.getElementById('imageInput').onchange = (e) => {
    const reader = new FileReader();
    reader.onload = (event) => {
        baseImg.src = event.target.result;
        baseImg.onload = update;
    };
    reader.readAsDataURL(e.target.files[0]);
};

// Carregar Excel/CSV com correção de caracteres especiais (UTF-8)
document.getElementById('excelInput').onchange = (e) => {
    const reader = new FileReader();
    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        // codepage: 65001 força a leitura em UTF-8 para evitar erro em acentos
        const wb = XLSX.read(data, { type: 'array', codepage: 65001 });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        dados = XLSX.utils.sheet_to_json(sheet);
        
        document.getElementById('counter').innerText = `${dados.length} registros encontrados`;
        document.getElementById('btnZip').disabled = false;
        update();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
};

// Ouvintes dos controles de ajuste
configIds.forEach(id => {
    document.getElementById(id).oninput = () => {
        if(id === 'yPos') document.getElementById('valY').innerText = document.getElementById(id).value;
        if(id === 'maxWidth') document.getElementById('valW').innerText = document.getElementById(id).value;
        update();
    };
});

function wrapText(text, x, y, maxWidth, lineHeight, context) {
    const words = text.split(' ');
    let line = '';
    let currentY = parseInt(y);
    
    for(let n = 0; n < words.length; n++) {
        let testLine = line + words[n] + ' ';
        if (context.measureText(testLine).width > maxWidth && n > 0) {
            context.fillText(line, x, currentY);
            line = words[n] + ' ';
            currentY += parseInt(lineHeight);
        } else { 
            line = testLine; 
        }
    }
    context.fillText(line, x, currentY);
}

function update() {
    if (!baseImg.src) return;
    canvas.width = baseImg.width; 
    canvas.height = baseImg.height;
    
    ctx.drawImage(baseImg, 0, 0);
    ctx.font = `${document.getElementById('fontSize').value}px "Montserrat", Arial`;
    ctx.fillStyle = "#000000"; 
    ctx.textAlign = "center"; 
    
    // Exemplo para o preview (com acentos para testar)
    const preview = getTexto("CONCEIÇÃO DA SILVA ARAÚJO", "123.456.789-00", "FORMAÇÃO DE TESTE COM AÇÕES", "40");
    
    wrapText(
        preview, 
        canvas.width / 2, 
        document.getElementById('yPos').value, 
        document.getElementById('maxWidth').value, 
        document.getElementById('lineHeight').value, 
        ctx
    );
}

// ==========================================
// 3. GERAÇÃO DO ARQUIVO ZIP
// ==========================================
document.getElementById('btnZip').onclick = async function() {
    const loader = document.getElementById('loader-overlay');
    if (loader) loader.style.display = 'flex';
    
    try {
        const zip = new JSZip();
        const tCanvas = document.createElement('canvas');
        const tCtx = tCanvas.getContext('2d');
        tCanvas.width = baseImg.width; 
        tCanvas.height = baseImg.height;

        for (const p of dados) {
            tCtx.clearRect(0, 0, tCanvas.width, tCanvas.height);
            tCtx.drawImage(baseImg, 0, 0);
            
            tCtx.font = `${document.getElementById('fontSize').value}px "Montserrat", Arial`;
            tCtx.fillStyle = "black"; 
            tCtx.textAlign = "center";

            // Captura flexível de colunas
            const nome = p["Nome Completo"] || p["Nome"] || p["NOME"] || Object.values(p)[0] || "Sem_Nome";
            const cpf = p["CPF"] || Object.values(p)[1] || "000.000.000-00";
            const curso = p["Formação"] || p["Curso"] || p["FORMAÇÃO"] || "Formação";
            const horas = p["Horas"] || p["Carga Horária"] || Object.values(p)[2] || "0";

            const textoFinal = getTexto(nome, cpf, curso, horas);

            wrapText(
                textoFinal, 
                tCanvas.width / 2, 
                document.getElementById('yPos').value, 
                document.getElementById('maxWidth').value, 
                document.getElementById('lineHeight').value, 
                tCtx
            );
            
            const imgData = tCanvas.toDataURL('image/png').split(',')[1];
            // Nome do arquivo sem acentos para evitar erro de download
            const fileName = `Certificado_${removerAcentos(nome)}.png`;
            zip.file(fileName, imgData, {base64: true});
        }

        const content = await zip.generateAsync({type:"blob"});
        saveAs(content, "Certificados_SEI.zip");
    } catch (error) {
        console.error("Erro ao gerar ZIP:", error);
        alert("Ocorreu um erro ao gerar os arquivos.");
    } finally {
        if (loader) loader.style.display = 'none';
    }
};