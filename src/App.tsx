import React, { useState, useEffect, useMemo, useRef, Suspense } from 'react';
import { 
  FileText, 
  Settings, 
  Download, 
  PlusCircle, 
  Trash2, 
  Info, 
  ArrowRight,
  CheckCircle2,
  Copy,
  FileType,
  Upload,
  Loader2,
  Book,
  Save,
  Search,
  Bold,
  Italic,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Heading1,
  Heading2,
  Type,
  Image as ImageIcon
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { Document, Packer, Paragraph, TextRun, AlignmentType, Header, Footer, ImageRun, TextWrappingType, TextWrappingSide, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, BorderStyle } from 'docx';
import { saveAs } from 'file-saver';
import { GoogleGenAI } from "@google/genai";
import mammoth from 'mammoth';

import 'react-quill-new/dist/quill.snow.css';
const ReactQuill = React.lazy(() => import('react-quill-new'));

// --- SVG Logo Component ---
const SimplesLogo = ({ className = "" }: { className?: string }) => {
  const [error, setError] = useState(false);
  if (error) {
    return <span className={`font-cinzel font-bold text-inherit block text-center ${className}`}>SIMPLES</span>;
  }
  return <img src="/logo.png" alt="Simples Logo" className={`object-contain ${className}`} onError={() => setError(true)} />;
};

// --- Types ---
interface Variable {
  key: string; // The text inside brackets, e.g., "NOME DA EMPRESA"
  value: string;
}

interface SavedTemplate {
  id: string;
  name: string;
  content: string;
  lastUsed: number;
}

// --- Utils ---
const extractVariables = (text: string): string[] => {
  // We look for text inside brackets [TEXT]
  const regex = /\[(.*?)\]/g;
  const matches = new Set<string>();
  let match;
  
  // Clean text from common HTML tags and decode &nbsp; to avoid [NOME&nbsp;DA&nbsp;EMPRESA]
  const cleanText = text.replace(/<[^>]*>?/gm, '').replace(/&nbsp;/gi, ' ');
  
  while ((match = regex.exec(cleanText)) !== null) {
    const cleanedVar = match[1].trim().replace(/\s+/g, ' ');
    if (cleanedVar) {
      matches.add(cleanedVar);
    }
  }
  return Array.from(matches);
};

const fillTemplate = (template: string, variables: Variable[]): string => {
  // Normalize &nbsp; to regular space in the template text before substitution
  let filled = template.replace(/&nbsp;/gi, ' ');
  variables.forEach((v) => {
    const escapedKey = v.key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(`\\[${escapedKey}\\]`, 'g');
    // For rich text, we just replace the markers
    filled = filled.replace(regex, v.value || `[${v.key}]`);
  });
  return filled;
};

// --- AI Service ---
const analyzeDocument = async (fileBase64: string, mimeType: string, originalText?: string) => {
  const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
  
  const prompt = `Analise este documento e retorne o texto integral dele, mas substituindo campos que claramente são variáveis (como nomes, CPFs, datas, endereços, nomes de empresas) por marcadores entre colchetes, como [NOME], [CPF], [DATA], [ENDEREÇO], [NOME DA EMPRESA]. 
  
  Se o documento já contiver marcadores entre colchetes, mantenha-os e apenas identifique se faltou algum campo importante para ser transformado em variável.
  
  Retorne APENAS o texto do modelo resultante. Nada de conversas, apenas o conteúdo do documento com os colchetes.`;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: [
        {
          parts: [
            { text: prompt },
            ...(mimeType === 'application/pdf' 
              ? [{ inlineData: { data: fileBase64, mimeType } }] 
              : [{ text: `Conteúdo extraído do arquivo: \n\n${originalText}` }])
          ]
        }
      ],
    });
    return response.text;
  } catch (err) {
    console.error("Erro AI:", err);
    throw err;
  }
};

// --- Constants: Default Templates ---
const DEFAULT_TEMPLATES: SavedTemplate[] = [
  {
    id: 'tpl-nda',
    name: 'NDA - Confidencialidade e Não Concorrência',
    content: `<h1 class="ql-align-center"><strong>ACORDO DE CONFIDENCIALIDADE, NÃO CONCORRÊNCIA E NÃO ALICIAMENTO</strong></h1><p><br></p><p>Pelo presente instrumento particular, de um lado [NOME DA EMPRESA], pessoa jurídica de direito privado, inscrita no CNPJ sob o nº [CNPJ DA EMPRESA], doravante denominada EMPRESA, e de outro lado, [NOME DO COLABORADOR], inscrito(a) no CPF sob o nº [CPF DO COLABORADOR], residente doravante denominado(a) PROFISSIONAL, celebram o presente acordo mediante as cláusulas e condições abaixo:</p><p><br></p><h2><strong>1. DA CONFIDENCIALIDADE E SIGILO (NDA)</strong></h2><p>O PROFISSIONAL obriga-se a manter o mais absoluto sigilo sobre toda e qualquer Informação Confidencial da EMPRESA a que tiver acesso em razão de seu vínculo, ressalvadas as hipóteses de autorização prévia e expressa, por escrito, por parte da EMPRESA.</p><p><br></p><p><strong>1.1. Abrangência:</strong> Entende-se por "Informação Confidencial" todo e qualquer dado técnico, operacional, comercial, jurídico ou financeiro, incluindo, mas não se limitando a: know-how, segredos de negócio, algoritmos, códigos-fonte, estratégias de marketing, precificação, listas de fornecedores, dados de faturamento e planos de expansão.</p><p><strong>1.2. Vigência:</strong> Esta obrigação de sigilo possui caráter irrevogável e irretratável, sobrevivendo ao término do vínculo por prazo indeterminado, ou enquanto referida informação não se tornar de domínio público por meios legítimos.</p><p><br></p><h2><strong>2. DO NÃO ALICIAMENTO (CLIENTES E TALENTOS)</strong></h2><p>Pelo período de 24 meses após o encerramento do vínculo, independentemente do motivo da rescisão, o PROFISSIONAL não poderá, direta ou indiretamente, praticar os atos abaixo descritos, salvo se houver acordo específico por escrito entre as Partes:</p><p><br></p><p><strong>2.1. Carteira de Clientes:</strong> Contatar, prospectar, desviar ou aceitar como cliente (seja via empresa própria, consultoria ou novo emprego) qualquer pessoa física ou jurídica que tenha sido cliente ou prospect da EMPRESA nos últimos 12 meses anteriores ao seu desligamento.</p><p><strong>2.2. Corpo Funcional:</strong> Aliciar, convidar ou contratar empregados, prestadores de serviço ou parceiros estratégicos da EMPRESA para compor novas estruturas de negócio fora da organização.</p><p><br></p><h2><strong>3. DO NÃO CONCORRÊNCIA (NON-COMPETE)</strong></h2><p>O PROFISSIONAL compromete-se a não exercer atividades que entrem em conflito direto com os interesses da EMPRESA, a menos que as Partes venham a pactuar de forma diversa por meio de termo aditivo ou autorização formal escrita da EMPRESA.</p><p><br></p><p><strong>3.1. Restrição:</strong> Pelo prazo de 12 meses, o profissional não poderá prestar serviços, possuir participação societária ou atuar como consultor para empresas que operem no mesmo segmento de mercado da EMPRESA.</p><p><strong>3.2. Limitação Geográfica:</strong> Esta restrição aplica-se a todo o território de atuação comercial da mesma.</p><p><br></p><h2><strong>4. PROPRIEDADE INTELECTUAL</strong></h2><p>Toda e qualquer invenção, melhoria, processo ou obra intelectual criada pelo PROFISSIONAL durante a vigência do contrato, que decorra da natureza de seu trabalho, pertencerá exclusivamente à EMPRESA, detendo esta todos os direitos de exploração comercial e propriedade.</p><p><br></p><h2><strong>5. DA CLÁUSULA PENAL E INDENIZAÇÕES</strong></h2><p>O descumprimento de qualquer obrigação prevista neste termo, sem a devida autorização ou acordo prévio por escrito, sujeitará o infrator:</p><p><br></p><p>a) Ao pagamento imediato de MULTA PENAL compensatória no valor de um salário mínimo vigente a época;</p><p>b) À apuração de PERDAS E DANOS suplementares, caso o prejuízo efetivo à EMPRESA supere o valor da multa estipulada;</p><p>c) À obtenção de medidas judiciais de natureza coercitiva (liminares) para cessação imediata da atividade infratora.</p><p><br></p><h2><strong>6. DISPOSIÇÕES GERAIS</strong></h2><p><strong>6.1. Flexibilização por Acordo:</strong> As restrições previstas neste instrumento poderão ser mitigadas, alteradas ou canceladas a qualquer tempo, desde que haja mútuo consentimento formalizado por escrito entre o PROFISSIONAL e a EMPRESA.</p><p><strong>6.2. Independência das Cláusulas:</strong> A eventual invalidade de uma cláusula não afetará a validade das demais.</p><p><strong>6.3. Foro:</strong> Fica eleito o Foro da Comarca de [CIDADE/UF] para dirimir quaisquer dúvidas oriundas deste termo.</p><p><br></p><p>[CIDADE/UF], [DATA].</p><p><br></p><p>__________________________________________</p><p><strong>[NOME DA EMPRESA]</strong></p><p><br></p><p>__________________________________________</p><p><strong>[NOME DO COLABORADOR]</strong></p>`,
    lastUsed: Date.now()
  },
  {
    id: 'tpl-banco-horas',
    name: 'Acordo Individual - Banco de Horas',
    content: `<h1 class="ql-align-center"><strong>INSTRUMENTO PARTICULAR DE ACORDO INDIVIDUAL PARA PRORROGAÇÃO E COMPENSAÇÃO DE JORNADA DE TRABALHO</strong></h1><p><br></p><p><strong>EMPREGADOR:</strong> [NOME DA EMPRESA], inscrita no CNPJ/MF sob o nº [CNPJ DA EMPRESA].</p><p><strong>EMPREGADO(A):</strong> [NOME DO COLABORADOR], portador(a) do CPF nº [CPF DO COLABORADOR].</p><p><br></p><p>As partes acima qualificadas celebram entre si o presente Acordo Individual, mediante as seguintes cláusulas:</p><p><br></p><h3><strong>CLÁUSULA PRIMEIRA – DO OBJETO</strong></h3><p>O presente acordo tem por objetivo instituir o regime de Prorrogação de Jornada cumulado com Compensação de Horas (Banco de Horas), permitindo que as horas trabalhadas além da jornada contratual sejam compensadas com folgas ou reduções de jornada, ou, pagas como horas extraordinárias.</p><p><br></p><h3><strong>CLÁUSULA SEGUNDA – DA SOLICITAÇÃO E DA VOLUNTARIEDADE (VIDA PESSOAL)</strong></h3><p>A prestação de horas extraordinárias reger-se-á pelos seguintes critérios de mútua anuência:</p><p><em>Iniciativa do Empregador:</em> O EMPREGADO apenas poderá realizar horas extraordinárias mediante solicitação expressa ou autorização prévia e formal do EMPREGADOR ou superior hierárquico.</p><p><em>Faculdade de Recusa:</em> Caso a convocação para jornada extraordinária interfira de forma comprovada ou relevante na vida pessoal e compromissos extraoficiais do EMPREGADO, fica a este facultado o direito de declinar do cumprimento da hora extra em questão, sem que isso configure insubordinação ou infração disciplinar.</p><p><br></p><h3><strong>CLÁUSULA TERCEIRA – DOS LIMITES E REGISTRO</strong></h3><p>A prorrogação da jornada não poderá exceder o limite legal de 02 (duas) horas diárias, totalizando o máximo de 10 (dez) horas de trabalho por dia.</p><p><em>Parágrafo Único:</em> A permanência nas dependências da empresa ou logado em sistemas remotos sem solicitação do EMPREGADOR será considerada tempo de natureza particular, não sendo computada como hora extra.</p><p><br></p><h3><strong>CLÁUSULA QUARTA – DA FORMA DE QUITAÇÃO (CRITÉRIO DO EMPREGADOR)</strong></h3><p>Fica a critério exclusivo do EMPREGADOR decidir se as horas excedentes serão destinadas ao Banco de Horas (compensação) ou se serão pagas como Horas Extras no contracheque.</p><p><em>Comunicação Prévia:</em> O EMPREGADOR deverá comunicar ao EMPREGADO, com antecedência, qual será a modalidade de quitação escolhida (se haverá folga compensatória ou pagamento).</p><p><br></p><h3><strong>CLÁUSULA QUINTA – DO REGIME DE COMPENSAÇÃO (BANCO DE HORAS)</strong></h3><p>Sendo adotada a compensação, observar-se-á:</p><p><em>Prazo:</em> A compensação das horas acumuladas deverá ocorrer no prazo máximo de 06 (seis) meses, conforme Art. 59, § 5º da CLT.</p><p><em>Gestão de Folgas:</em> A definição das datas e períodos de folga compensatória será de conveniência do EMPREGADOR.</p><p><br></p><h3><strong>CLÁUSULA SEXTA – DA REMUNERAÇÃO DAS HORAS NÃO COMPENSADAS</strong></h3><p>Caso as horas excedentes não sejam compensadas dentro do prazo de 06 meses, o EMPREGADOR efetuará o pagamento destas como horas extras, com o adicional mínimo de 50% (cinquenta por cento) ou o percentual mais benéfico previsto em Convenção Coletiva de Trabalho (CCT).</p><p><br></p><h3><strong>CLÁUSULA SÉTIMA – DO SALDO NEGATIVO</strong></h3><p>Eventuais débitos de horas do EMPREGADO (atrasos ou saídas antecipadas autorizadas) poderão ser compensados com horas positivas ou, ao final do período de 06 meses/rescisão, serem descontados em folha de pagamento.</p><p><br></p><h3><strong>CLÁUSULA OITAVA – DA RESCISÃO CONTRATUAL</strong></h3><p>Na hipótese de rescisão do contrato de trabalho:</p><p><em>Saldo Positivo:</em> As horas não compensadas serão pagas como extras nas verbas rescisórias.</p><p><em>Saldo Negativo:</em> As horas não trabalhadas poderão ser descontadas das verbas rescisórias, conforme limite legal.</p><p><br></p><h3><strong>CLÁUSULA NONA – VIGÊNCIA E FORO</strong></h3><p>Este acordo tem validade por prazo indeterminado, podendo ser revisto ou aditado caso surjam novas Normas Coletivas que se sobreponham a estas condições. As partes elegem o foro da Comarca de [CIDADE/UF] para dirimir controvérsias.</p><p><br></p><p>[CIDADE/UF], [DATA].</p><p><br></p><p>___________________________________</p><p><strong>EMPREGADOR</strong></p><p><br></p><p>___________________________________</p><p><strong>EMPREGADO</strong></p>`,
    lastUsed: Date.now()
  },
  {
    id: 'tpl-etica-digital',
    name: 'Termo de Ética e Ferramentas Digitais',
    content: `<h1 class="ql-align-center"><strong>TERMO DE COMPROMISSO: ÉTICA, CONDUTA PROFISSIONAL E USO DE FERRAMENTAS DIGITAIS</strong></h1><p><br></p><p><strong>EMPREGADOR:</strong> [NOME DA EMPRESA], inscrito no CNPJ sob o nº [CNPJ DA EMPRESA].</p><p><strong>EMPREGADO:</strong> [NOME DO COLABORADOR], inscrito no CPF sob o nº [CPF DO COLABORADOR].</p><p><br></p><h3><strong>1. DOS PADRÕES ÉTICOS E COMPORTAMENTAIS</strong></h3><p>1.1. O EMPREGADO compromete-se a manter uma conduta pautada pela urbanidade, probidade e sigilo profissional, zelando pela reputação institucional da EMPRESA perante clientes, parceiros e órgãos públicos. </p><p>1.2. Constituem faltas graves, sujeitas às sanções do Art. 482 da CLT, sem prejuízo de responsabilidade civil:</p><p>a) Violação de Sigilo: Revelar métodos, processos, dados de clientes ou estratégias internas a terceiros.</p><p>b) Incontinência de Conduta: Comportamentos inadequados, assédio (moral ou sexual) ou linguagem incompatível com o ambiente corporativo.</p><p>c) Desídia: Negligência técnica, atrasos injustificados em prazos processuais ou falta de zelo com o patrimônio da empresa.</p><p>d) Ato Lesivo à Honra: Críticas públicas desairosas à empresa ou clientes em redes sociais ou canais de comunicação.</p><p><br></p><h3><strong>2. DA POLÍTICA DE COMUNICAÇÃO DIGITAL E FERRAMENTAS</strong></h3><p>2.1. Finalidade: O uso de aplicativos de mensagens (WhatsApp, Telegram, Slack), e-mails e sistemas internos destina-se exclusivamente ao desempenho das funções laborais, visando a eficiência e a agilidade técnica. </p><p>2.2. Direito à Desconexão: Fica estabelecido que o envio de comunicações por parte da EMPRESA ou de clientes fora da jornada contratual de trabalho não implica obrigatoriedade de visualização ou resposta imediata, não configurando tempo à disposição, prontidão ou sobreaviso.</p><p>Parágrafo Único: A regra acima prevalece, salvo em casos de acordo prévio e específico entre as partes, regimes de escala formalmente estabelecidos ou situações de urgência excepcional que demandem intervenção imediata, respeitando-se sempre as compensações legais. </p><p>2.3. Propriedade dos Dados: Todo o histórico de mensagens, arquivos e comunicações realizadas em contextos profissionais ou por meio de contas vinculadas à EMPRESA é de propriedade exclusiva do EMPREGADOR. É terminantemente proibida a exclusão de históricos ou arquivos sem autorização prévia, visando garantir a rastreabilidade e a segurança jurídica das operações.</p><p><br></p><h3><strong>3. DO MONITORAMENTO E PRIVACIDADE</strong></h3><p>3.1. O EMPREGADO declara ciência de que computadores, contas de e-mail corporativo, aparelhos celulares fornecidos pela empresa e sistemas de gestão são ferramentas de trabalho. </p><p>3.2. A EMPRESA reserva-se o direito de auditar e monitorar o uso destas ferramentas para fins de segurança da informação, conformidade com a LGPD (Lei Geral de Proteção de Dados) e verificação de produtividade, não havendo expectativa de privacidade sobre atos praticados no exercício da função.</p><p><br></p><h3><strong>4. PROTEÇÃO DE DADOS (LGPD)</strong></h3><p>4.1. O EMPREGADO obriga-se a tratar os dados pessoais a que tiver acesso em estrita observância à Lei nº 13.709/2018, utilizando-os apenas para as finalidades determinadas pelo EMPREGADOR e adotando medidas para evitar acessos não autorizados ou vazamentos.</p><p><br></p><h3><strong>5. DAS PENALIDADES</strong></h3><p>5.1. A inobservância de qualquer cláusula deste termo facultará à EMPRESA a aplicação de medidas disciplinares, que podem variar de advertência e suspensão até a rescisão do contrato de trabalho por justa causa, conforme a gravidade da infração e o histórico do colaborador.</p><p><br></p><p>[CIDADE/UF], [DATA].</p><p><br></p><p>________________________</p><p><strong>[NOME DO COLABORADOR]</strong></p><p><br></p><p>________________________</p><p><strong>[NOME DA EMPRESA]</strong></p>`,
    lastUsed: Date.now()
  },
  {
    id: 'tpl-imagem',
    name: 'Autorização de Uso de Imagem e Voz',
    content: `<h1 class="ql-align-center"><strong>TERMO DE CONSENTIMENTO E AUTORIZAÇÃO DE USO DE IMAGEM, VOZ E NOME</strong></h1><p><br></p><p><strong>EMPREGADOR:</strong> [NOME DA EMPRESA], pessoa jurídica de direito privado, inscrita no CNPJ sob o nº [CNPJ DA EMPRESA]</p><p><strong>EMPREGADO:</strong> [NOME DO COLABORADOR], inscrito no CPF sob o nº [CPF DO COLABORADOR]</p><p><br></p><p>Pelo presente instrumento, o EMPREGADO autoriza o EMPREGADOR, de forma livre, informada e inequívoca, a utilizar sua imagem, voz e nome, em conformidade com as cláusulas abaixo:</p><p><br></p><h3><strong>1. OBJETO E FINALIDADE:</strong></h3><p>A presente autorização destina-se ao uso da imagem (estática ou em movimento), voz e nome do EMPREGADO para fins de publicidade e divulgação de cultura corporativa, podendo ser utilizada em:</p><p><em>Mídias Digitais:</em> Redes sociais (Instagram, LinkedIn, Facebook, YouTube, TikTok, entre outras), site institucional, blogs e anúncios digitais.</p><p><em>Comunicação Interna:</em> Materiais de treinamento, murais físicos ou digitais e comunicados.</p><p><em>Materiais Impressos e Publicidade:</em> Folders, catálogos, outdoors, apresentações comerciais, revistas e jornais.</p><p><br></p><h3><strong>2. ABRANGÊNCIA E TERRITÓRIO:</strong></h3><p>A autorização é concedida em caráter global (Brasil e exterior), permitindo que o EMPREGADOR realize edições, cortes, fixações e reproduções do material, desde que preservada a honra e a imagem pública do EMPREGADO.</p><p><br></p><h3><strong>3. GRATUIDADE:</strong></h3><p>O EMPREGADO declara que a presente autorização é concedida de forma totalmente gratuita. O uso da imagem, voz e nome não gera direito a qualquer tipo de remuneração extra, "cachê", indenização ou participação financeira, a não ser que acordada diretamente com o empregador previamente, sendo o conteúdo produzido considerado parte da relação profissional estabelecida.</p><p><br></p><h3><strong>4. PRAZO E REVOGAÇÃO (DIREITO AO ARREPENDIMENTO):</strong></h3><p>A autorização é válida por prazo indeterminado, permanecendo vigente inclusive após o término do contrato de trabalho, observadas as seguintes condições:</p><p><em>Direito de Revogação:</em> O EMPREGADO poderá, a qualquer tempo, solicitar a revogação desta autorização mediante comunicação escrita ao setor de Recursos Humanos da empresa.</p><p><em>Efeito Non Retroativo:</em> Em caso de revogação, o EMPREGADOR interromperá a utilização do material em novas produções ou campanhas. Todavia, a empresa não possui obrigação de remover, apagar ou recolher materiais, publicações, vídeos ou impressos já executados, publicados ou distribuídos anteriormente à data da revogação.</p><p><br></p><h3><strong>5. PROTEÇÃO DE DADOS (LGPD):</strong></h3><p>O EMPREGADOR, na qualidade de Controlador de Dados, compromete-se a tratar os dados biovocais e de imagem do EMPREGADO em estrita observância à LGPD, garantindo que o tratamento seja limitado às finalidades institucionais aqui descritas, adotando medidas de segurança para proteger tais informações.</p><p><br></p><h3><strong>6. DISPOSIÇÕES GERAIS:</strong></h3><p>O EMPREGADO declara ter lido e compreendido todos os termos deste documento, estando de pleno acordo com a utilização de sua imagem e voz conforme aqui estipulado, renunciando a qualquer autorização prévia para edições que não alterem o contexto profissional do material.</p><p><br></p><p>[CIDADE/UF], [DATA].</p><p><br></p><p>_____________________________</p><p><strong>EMPREGADOR</strong></p><p><br></p><p>_____________________________</p><p><strong>EMPREGADO</strong></p>`,
    lastUsed: Date.now()
  }
];

// --- Components ---

export default function App() {
  const [template, setTemplate] = useState<string>(
    `<h1 class="ql-align-center"><strong>TERMO DE RESPONSABILIDADE</strong></h1><p><br></p><p><strong>NOME DA EMPRESA:</strong> [NOME DA EMPRESA]</p><p><strong>CNPJ:</strong> [CNPJ]</p><p><br></p><p>Através deste documento, confirmamos que o(a) Sr(a). <strong>[NOME DO COLABORADOR]</strong>, portador(a) do CPF [CPF], residente e domiciliado(a) em [ENDEREÇO], recebeu os equipamentos necessários para a execução de suas atividades.</p><p><br></p><p><strong>Data:</strong> [DATA]</p><p><br></p><p><strong>Assinatura:</strong> __________________________</p>`
  );
  
  const [variables, setVariables] = useState<Variable[]>([]);
  const [activeStep, setActiveStep] = useState<'editor' | 'filler' | 'preview'>('editor');
  const [copySuccess, setCopySuccess] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [savedTemplates, setSavedTemplates] = useState<SavedTemplate[]>([]);
  const [templateName, setTemplateName] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Quill Modules Configuration
  const quillModules = useMemo(() => ({
    toolbar: [
      [{ header: [1, 2, 3, false] }],
      ['bold', 'italic', 'underline', 'strike'],
      [{ list: 'ordered' }, { list: 'bullet' }],
      [{ align: [] }],
      ['clean'],
    ],
  }), []);

  const filteredTemplates = useMemo(() => {
    return savedTemplates.filter(t => 
      t.name.toLowerCase().includes(searchQuery.toLowerCase())
    );
  }, [savedTemplates, searchQuery]);

  // Load templates from LocalStorage
  useEffect(() => {
    const stored = localStorage.getItem('auto_termos_templates');
    if (stored) {
      setSavedTemplates(JSON.parse(stored));
    }
  }, []);

  // Save templates to LocalStorage
  const saveToLibrary = () => {
    const name = templateName.trim() || `Modelo ${savedTemplates.length + 1}`;
    const newTemplate: SavedTemplate = {
      id: crypto.randomUUID(),
      name,
      content: template,
      lastUsed: Date.now()
    };
    const updated = [newTemplate, ...savedTemplates];
    setSavedTemplates(updated);
    localStorage.setItem('auto_termos_templates', JSON.stringify(updated));
    setTemplateName("");
  };

  const deleteFromLibrary = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const updated = savedTemplates.filter(t => t.id !== id);
    setSavedTemplates(updated);
    localStorage.setItem('auto_termos_templates', JSON.stringify(updated));
  };

  const selectTemplate = (t: SavedTemplate) => {
    setTemplate(t.content);
    // Update last used
    const updated = savedTemplates.map(s => s.id === t.id ? { ...s, lastUsed: Date.now() } : s);
    setSavedTemplates(updated);
    localStorage.setItem('auto_termos_templates', JSON.stringify(updated));
  };

  // Sync variables with template
  useEffect(() => {
    const keys = extractVariables(template);
    setVariables(prev => {
      return keys.map(key => {
        const existing = prev.find(p => p.key === key);
        return { key, value: existing ? existing.value : '' };
      });
    });
  }, [template]);

  const filledText = useMemo(() => fillTemplate(template, variables), [template, variables]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsAnalyzing(true);
    try {
      const reader = new FileReader();
      
      if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        reader.onload = async (event) => {
          const arrayBuffer = event.target?.result as ArrayBuffer;
          const result = await mammoth.extractRawText({ arrayBuffer });
          const processed = await analyzeDocument("", "text/plain", result.value);
          if (processed) setTemplate(processed);
          setIsAnalyzing(false);
        };
        reader.readAsArrayBuffer(file);
      } else if (file.type === 'application/pdf') {
        reader.onload = async (event) => {
          const base64 = (event.target?.result as string).split(',')[1];
          const processed = await analyzeDocument(base64, "application/pdf");
          if (processed) setTemplate(processed);
          setIsAnalyzing(false);
        };
        reader.readAsDataURL(file);
      } else {
        alert("Por favor, envie um arquivo .docx ou .pdf");
        setIsAnalyzing(false);
      }
    } catch (err) {
      console.error(err);
      alert("Erro ao processar arquivo com IA. Verifique sua conexão ou tente novamente.");
      setIsAnalyzing(false);
    }
  };

  const printRef = useRef<HTMLDivElement>(null);

  const handleDownloadDocx = async () => {
    // Basic conversion from HTML to Docx (mammoth/quill style)
    
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = filledText;
    
    const docChildren: Paragraph[] = [];
    
    // Instead of just childNodes, let's query all block level elements that contain the text
    const elements = tempDiv.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li');
    
    if (elements.length > 0) {
      elements.forEach(el => {
        const isCenter = (el as HTMLElement).style?.textAlign === 'center' || el.classList?.contains('ql-align-center');
        const isRight = (el as HTMLElement).style?.textAlign === 'right' || el.classList?.contains('ql-align-right');
        const isJustify = (el as HTMLElement).style?.textAlign === 'justify' || el.classList?.contains('ql-align-justify');
        
        const alignment = isCenter ? AlignmentType.CENTER : 
                          isRight ? AlignmentType.RIGHT : 
                          isJustify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT;
        
        const textOptions: any = {
          text: el.textContent || '',
          size: el.nodeName === 'H1' ? 32 : el.nodeName === 'H2' ? 28 : el.nodeName === 'H3' ? 24 : 22
        };

        if (el.nodeName.startsWith('H') || el.innerHTML?.includes('<strong>') || el.innerHTML?.includes('<b>')) {
          textOptions.bold = true;
        }
        if (el.innerHTML?.includes('<em>') || el.innerHTML?.includes('<i>')) {
          textOptions.italics = true;
        }

        const paragraphOptions: any = {
          children: [new TextRun(textOptions)],
          spacing: { after: 200 },
          alignment
        };

        if (el.nodeName === 'LI') {
          paragraphOptions.bullet = { level: 0 };
        }

        docChildren.push(new Paragraph(paragraphOptions));
      });
    } else {
      docChildren.push(new Paragraph({ children: [new TextRun(filledText.replace(/<[^>]*>?/gm, ''))] }));
    }

    let headers: any = undefined;
    let footers: any = undefined;

    let bgFileBytes: Uint8Array | null = null;
    try {
      const resp = await fetch('/a.png');
      if (resp.ok) {
        const arrayBuffer = await resp.arrayBuffer();
        bgFileBytes = new Uint8Array(arrayBuffer);
      }
    } catch (e) {
      console.warn("Could not load a.png");
    }

    if (bgFileBytes) {
      const imageWidth = 794; 
      const imageHeight = 1123; 

      headers = {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new ImageRun({
                  data: bgFileBytes,
                  transformation: {
                    width: imageWidth,
                    height: imageHeight,
                  },
                  floating: {
                    horizontalPosition: {
                      relative: HorizontalPositionRelativeFrom.PAGE,
                      offset: 0,
                    },
                    verticalPosition: {
                      relative: VerticalPositionRelativeFrom.PAGE,
                      offset: 0,
                    },
                    behindDocument: true,
                    wrap: {
                      type: TextWrappingType.NONE,
                      side: TextWrappingSide.BOTH_SIDES,
                    },
                  },
                }),
              ],
            }),
          ],
        })
      };
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 4500,     // ~80mm to clear the letterhead header with breathing room
              bottom: 2800,  // ~50mm to clear the letterhead footer with breathing room
              left: 1400,    // ~25mm to clear the border frame
              right: 1400,   // ~25mm to clear the border frame
            },
          },
        },
        headers: headers,
        footers: footers,
        children: docChildren,
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `Termo_Preenchido_${new Date().getTime()}.docx`);
  };

  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false);

  const handleDownloadPdf = async () => {
    // Generate PDF via backend: sends HTML to server, which creates DOCX
    // and converts to PDF using LibreOffice for perfect formatting.
    setIsGeneratingPdf(true);
    try {
      const response = await fetch('/api/generate-pdf', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filledHtml: filledText }),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.details || `Erro ${response.status}`);
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Termo_Preenchido_${Date.now()}.pdf`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: any) {
      console.error('Erro ao gerar PDF:', err);
      alert(`Erro ao gerar PDF: ${err.message}\n\nVerifique se o servidor backend está rodando (npm run server).`);
    } finally {
      setIsGeneratingPdf(false);
    }
  };

  const handleCopyClipboard = () => {
    // Copy plain text for clipboard
    const plainText = filledText.replace(/<[^>]*>?/gm, '\n').replace(/\n\n+/g, '\n\n');
    navigator.clipboard.writeText(plainText);
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  return (
    <div className="min-h-screen font-sans selection:bg-bg-bordo text-text-creme">
      {/* Header */}
      <header className="border-b border-bg-bordo bg-bg-ebano/95 backdrop-blur-md sticky top-0 z-20">
        <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="h-8 flex items-center justify-center text-ouro-imperial">
               <SimplesLogo className="h-full w-auto object-contain" />
            </div>
            <h1 className="text-xl font-cinzel font-bold tracking-tight text-ouro-imperial uppercase mt-1"><span className="text-text-terra font-normal text-sm md:text-lg hidden md:inline ml-1">- AutoTermos</span></h1>
          </div>
          
          <nav className="hidden md:flex items-center gap-8">
            <button 
              onClick={() => setActiveStep('editor')}
              className={`text-sm font-cinzel font-bold tracking-widest uppercase transition-colors ${activeStep === 'editor' ? 'text-ouro-imperial' : 'text-text-terra hover:text-text-creme'}`}
            >
              1. Modelo
            </button>
            <button 
              onClick={() => setActiveStep('filler')}
              className={`text-sm font-cinzel font-bold tracking-widest uppercase transition-colors ${activeStep === 'filler' ? 'text-ouro-imperial' : 'text-text-terra hover:text-text-creme'}`}
            >
              2. Preencher
            </button>
            <button 
              onClick={() => setActiveStep('preview')}
              className={`text-sm font-cinzel font-bold tracking-widest uppercase transition-colors ${activeStep === 'preview' ? 'text-ouro-imperial' : 'text-text-terra hover:text-text-creme'}`}
            >
              3. Concluído
            </button>
          </nav>

          <div className="flex items-center gap-2">
            <button className="p-2 text-text-terra hover:text-ouro-imperial transition-colors">
              <Settings size={20} />
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 md:py-12">
        <AnimatePresence mode="wait">
          {/* STEP 1: EDITOR */}
          {activeStep === 'editor' && (
            <motion.div 
              key="editor"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="grid grid-cols-1 lg:grid-cols-12 gap-8"
            >
              {/* Library Sidebar */}
              <div className="lg:col-span-3 space-y-6">
                <div className="bg-bg-noite rounded-xl border border-bg-bordo shadow-inner overflow-hidden min-h-[500px] flex flex-col">
                  <div className="p-4 border-b border-bg-bordo space-y-3 bg-bg-bordo-profundo/30">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2 text-ouro-imperial font-cinzel font-bold text-sm uppercase tracking-wider">
                        <Book size={16} className="text-ouro-imperial" />
                        Biblioteca
                      </div>
                    </div>
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-text-terra" size={14} />
                      <input 
                        type="text"
                        placeholder="Buscar modelo..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="w-full pl-9 pr-3 py-2 bg-bg-ebano border border-bg-bordo rounded-lg text-xs text-text-creme outline-none focus:ring-1 focus:ring-ouro-imperial transition-all font-medium shadow-sm placeholder:text-text-terra"
                      />
                    </div>
                  </div>
                  
                  <div className="flex-1 overflow-y-auto p-2 space-y-4 custom-scrollbar">
                    {/* Default templates section */}
                    {!searchQuery && (
                      <div className="space-y-1">
                        <div className="px-3 py-1 text-[10px] font-cinzel font-bold text-text-areia uppercase tracking-widest">Sugeridos</div>
                        {DEFAULT_TEMPLATES.map(t => (
                          <button
                            key={t.id}
                            onClick={() => selectTemplate(t)}
                            className={`w-full text-left p-3 rounded-lg hover:bg-bg-bordo-profundo transition-colors flex flex-col border border-transparent ${template === t.content ? 'bg-bg-bordo border-ouro-escuro/30' : ''}`}
                          >
                            <div className="text-sm font-semibold text-text-creme truncate">{t.name}</div>
                            <div className="text-[10px] text-ouro-escuro font-bold uppercase tracking-widest mt-0.5">Padrão do Sistema</div>
                          </button>
                        ))}
                      </div>
                    )}

                    {/* User templates section */}
                    <div className="space-y-1">
                      <div className="px-3 py-1 text-[10px] font-cinzel font-bold text-text-areia uppercase tracking-widest">
                        {searchQuery ? "Resultados" : "Seus Modelos"}
                      </div>
                      {filteredTemplates.length === 0 && !searchQuery ? (
                        <div className="p-4 text-center text-text-terra text-[10px] italic">
                          Nenhum modelo personalizado.
                        </div>
                      ) : (
                        filteredTemplates.map(t => (
                          <div
                            key={t.id}
                            className="group relative"
                          >
                            <button
                              onClick={() => selectTemplate(t)}
                              className={`w-full text-left p-3 pr-10 rounded-lg hover:bg-bg-bordo-profundo transition-colors flex flex-col border border-transparent ${template === t.content ? 'bg-bg-bordo border-ouro-escuro/30' : ''}`}
                            >
                              <div className="text-sm font-semibold text-text-creme truncate pr-2">{t.name}</div>
                              <div className="text-[10px] text-text-terra uppercase tracking-widest mt-0.5">
                                Criado em {new Date(t.lastUsed).toLocaleDateString()}
                              </div>
                            </button>
                            <button
                              onClick={(e) => deleteFromLibrary(t.id, e)}
                              className="absolute right-2 top-1/2 -translate-y-1/2 p-2 text-gray-300 hover:text-red-500 hover:bg-red-50 rounded-md transition-all active:scale-95"
                              title="Deletar este modelo"
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        ))
                      )}
                      
                      {searchQuery && filteredTemplates.length === 0 && (
                         <div className="p-8 text-center text-text-terra text-xs italic">
                            Nenhum modelo encontrado.
                         </div>
                      )}
                    </div>
                  </div>
                </div>

                <div className="bg-bg-bordo-profundo border border-bg-bordo p-6 rounded-xl space-y-4 relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-32 h-32 bg-ouro-imperial/5 rounded-full blur-3xl" />
                  <h4 className="text-xs font-cinzel font-bold text-ouro-claro uppercase tracking-widest relative z-10">Importar com IA</h4>
                  <p className="text-xs text-text-creme/80 leading-relaxed relative z-10">Suba um Word ou PDF e nossa IA identificará as variáveis automaticamente.</p>
                  
                  <input 
                    type="file" 
                    ref={fileInputRef}
                    onChange={handleFileUpload}
                    accept=".pdf,.docx"
                    className="hidden"
                  />
                  
                  <button 
                    onClick={() => fileInputRef.current?.click()}
                    disabled={isAnalyzing}
                    className="w-full py-2.5 px-4 bg-bg-ebano border border-ouro-escuro/50 text-ouro-claro rounded-lg text-sm font-semibold hover:bg-bg-bordo transition-all flex items-center justify-center gap-2 shadow-sm disabled:opacity-50 relative z-10"
                  >
                    {isAnalyzing ? <Loader2 size={16} className="animate-spin" /> : <Upload size={16} />}
                    {isAnalyzing ? 'Analisando...' : 'Subir Arquivo'}
                  </button>
                </div>
              </div>

              {/* Main Content Area */}
              <div className="lg:col-span-6 space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex-1 max-w-xs">
                    <input 
                      type="text" 
                      value={templateName}
                      onChange={(e) => setTemplateName(e.target.value)}
                      placeholder="Nome do Modelo (Opcional)"
                      className="w-full bg-transparent border-b border-bg-bordo focus:border-ouro-imperial outline-none px-1 py-1 text-sm font-medium text-text-creme placeholder:text-text-terra"
                    />
                  </div>
                  <button 
                    onClick={saveToLibrary}
                    className="flex items-center gap-2 text-xs font-cinzel font-bold text-ouro-imperial hover:bg-bg-bordo px-3 py-1.5 rounded-lg transition-colors uppercase tracking-wider border border-transparent hover:border-ouro-escuro/30"
                  >
                    <Save size={14} />
                    Salvar na Biblioteca
                  </button>
                </div>
                
                <div className="relative group shadow-sm bg-champagne rounded-xl border border-ouro-escuro/30 focus-within:ring-2 focus-within:ring-ouro-imperial focus-within:border-transparent transition-all overflow-hidden text-black">
                  <Suspense fallback={
                    <div className="w-full h-[600px] flex items-center justify-center bg-bg-noite">
                      <Loader2 className="animate-spin text-ouro-imperial" size={32} />
                    </div>
                  }>
                    <ReactQuill
                      theme="snow"
                      value={template}
                      onChange={setTemplate}
                      modules={quillModules}
                      placeholder="Ex: [NOME] compareceu no dia [DATA]..."
                      className="h-[550px] mb-12"
                    />
                  </Suspense>
                  
                  {isAnalyzing && (
                    <div className="absolute inset-0 bg-bg-ebano/80 backdrop-blur-sm flex flex-col items-center justify-center gap-4 rounded-xl z-10 text-ouro-imperial text-center p-8">
                      <div className="relative">
                        <Loader2 size={40} className="animate-spin" />
                        <motion.div 
                          className="absolute inset-0 flex items-center justify-center"
                          animate={{ scale: [0.8, 1.2, 0.8] }}
                          transition={{ repeat: Infinity, duration: 2 }}
                        >
                          <Search size={16} />
                        </motion.div>
                      </div>
                      <span className="font-cinzel font-bold text-sm tracking-widest uppercase mt-4">IA Analisando Documento</span>
                      <p className="text-xs text-text-areia mt-2">Identificando variáveis e estruturando o modelo...</p>
                    </div>
                  )}
                </div>
              </div>

              {/* Action Sidebar */}
              <div className="lg:col-span-3 space-y-6">
                <div className="bg-bg-noite p-6 rounded-xl border border-bg-bordo shadow-inner space-y-4 sticky top-24 relative overflow-hidden">
                  <div className="absolute -right-10 -top-10 w-40 h-40 bg-bg-bordo/10 blur-3xl rounded-full pointer-events-none" />
                  <div className="flex items-center gap-2 text-ouro-claro relative z-10">
                    <Info size={18} />
                    <h3 className="font-cinzel font-bold uppercase tracking-widest text-xs">Instruções</h3>
                  </div>
                  <ul className="space-y-3 text-sm text-text-creme/80 relative z-10">
                    <li className="flex gap-2">
                      <span className="text-ouro-imperial font-bold">•</span>
                      <span>Tudo entre <b>[colchetes]</b> virará uma pergunta automática.</span>
                    </li>
                    <li className="flex gap-2">
                      <span className="text-ouro-imperial font-bold">•</span>
                      <span>Você pode repetir o mesmo marcador várias vezes.</span>
                    </li>
                  </ul>
                  
                  <div className="pt-4 space-y-2 relative z-10">
                    <h4 className="text-[10px] font-cinzel font-bold text-text-terra uppercase tracking-widest">Resumo</h4>
                    <div className="flex items-center justify-between text-sm">
                      <span className="text-text-areia">Variáveis únicas</span>
                      <span className="font-bold text-bg-ebano bg-ouro-claro px-2 py-0.5 rounded">{variables.length}</span>
                    </div>
                  </div>

                  <button 
                    onClick={() => setActiveStep('filler')}
                    disabled={variables.length === 0}
                    className="w-full mt-6 py-4 px-4 bg-gradient-to-r from-ouro-escuro to-ouro-imperial hover:from-ouro-imperial hover:to-ouro-claro disabled:from-bg-bordo disabled:to-bg-bordo disabled:text-text-terra disabled:border-bg-bordo disabled:cursor-not-allowed text-bg-ebano rounded-xl font-cinzel font-bold text-sm tracking-wider uppercase transition-all flex items-center justify-center gap-3 shadow-lg shadow-ouro-imperial/20 hover:shadow-ouro-imperial/40 relative z-10"
                  >
                    Começar Preenchimento
                    <ArrowRight size={18} />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 2: FILLER */}
          {activeStep === 'filler' && (
            <motion.div 
              key="filler"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="max-w-3xl mx-auto space-y-8"
            >
              <div className="text-center space-y-2">
                <h2 className="text-3xl font-cinzel font-bold text-ouro-imperial uppercase tracking-wide">Alimentando o Termo</h2>
                <p className="text-text-areia font-serif italic text-lg">Insira os dados abaixo para personalizar o documento.</p>
              </div>

              <div className="bg-bg-noite border-t-2 border-t-ouro-imperial border-x border-b border-x-bg-bordo border-b-bg-bordo rounded-2xl shadow-2xl shadow-black/50 overflow-hidden relative">
                <div className="absolute top-0 right-0 w-64 h-64 bg-bg-bordo-profundo/40 blur-3xl rounded-full pointer-events-none" />
                <div className="p-8 space-y-6 max-h-[600px] overflow-y-auto custom-scrollbar relative z-10">
                  {variables.map((variable, idx) => (
                    <div key={variable.key} className="space-y-1.5 group">
                      <label htmlFor={`var-${idx}`} className="block text-xs font-cinzel font-bold text-text-terra uppercase tracking-widest px-1 group-focus-within:text-ouro-claro transition-colors drop-shadow-sm">
                        {variable.key}
                      </label>
                      <input
                        id={`var-${idx}`}
                        type="text"
                        value={variable.value}
                        onChange={(e) => {
                          const newVars = [...variables];
                          newVars[idx].value = e.target.value;
                          setVariables(newVars);
                        }}
                        placeholder={`Digite o ${variable.key.toLowerCase()}...`}
                        className="w-full px-4 py-4 bg-bg-ebano border border-bg-bordo rounded-xl focus:bg-bg-ebano focus:ring-1 focus:ring-ouro-imperial focus:border-ouro-imperial outline-none transition-all placeholder:text-bg-bordo-profundo text-lg sm:text-base font-serif text-text-creme"
                      />
                    </div>
                  ))}

                  {variables.length === 0 && (
                     <div className="text-center py-8 space-y-4">
                        <Trash2 className="mx-auto text-bg-bordo" size={48} />
                        <p className="text-text-terra text-sm">Nenhuma variável detectada.</p>
                        <button 
                          onClick={() => setActiveStep('editor')}
                          className="text-ouro-imperial font-cinzel tracking-widest uppercase text-xs hover:text-ouro-claro"
                        >
                          Voltar ao Modelo
                        </button>
                     </div>
                  )}
                </div>

                <div className="bg-bg-bordo-profundo/50 p-8 flex flex-col sm:flex-row gap-4 border-t border-bg-bordo relative z-10">
                  <button 
                    onClick={() => setActiveStep('editor')}
                    className="flex-1 py-4 px-6 bg-transparent border border-ouro-escuro/50 text-text-creme rounded-xl font-cinzel tracking-wider uppercase font-bold hover:bg-ouro-imperial/10 transition-colors"
                  >
                    Editar Modelo
                  </button>
                  <button 
                    onClick={() => setActiveStep('preview')}
                    className="flex-[3] py-4 px-6 bg-gradient-to-r from-ouro-escuro to-ouro-imperial text-bg-ebano rounded-xl font-cinzel tracking-wider uppercase font-bold hover:from-ouro-imperial hover:to-ouro-claro transition-all shadow-xl shadow-ouro-imperial/20 flex items-center justify-center gap-3 transform hover:-translate-y-0.5 active:translate-y-0"
                  >
                    Gerar Documento
                    <CheckCircle2 size={20} />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 3: PREVIEW / EXPORT */}
          {activeStep === 'preview' && (
            <motion.div 
              key="preview"
              initial={{ opacity: 0, scale: 0.98 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.98 }}
              className="grid grid-cols-1 lg:grid-cols-4 gap-8"
            >
              <div className="lg:col-span-1 space-y-4 order-2 lg:order-1">
                <h3 className="text-sm font-cinzel font-bold text-ouro-imperial uppercase tracking-widest">Opções de Saída</h3>
                <div className="space-y-3">
                  <button 
                    onClick={handleDownloadDocx}
                    className="w-full flex items-center gap-4 p-5 bg-bg-noite border border-bg-bordo rounded-2xl hover:border-ouro-imperial hover:shadow-lg transition-all text-left shadow-sm group"
                  >
                    <div className="p-3 bg-bg-bordo-profundo rounded-xl group-hover:bg-bg-bordo transition-colors">
                      <FileType className="text-ouro-imperial" size={24} />
                    </div>
                    <div>
                      <div className="font-cinzel font-bold text-base text-text-creme">Baixar Word</div>
                      <div className="text-[10px] text-text-terra font-bold uppercase tracking-widest">Documento Editável</div>
                    </div>
                  </button>

                  <button 
                    onClick={handleCopyClipboard}
                    className="w-full flex items-center gap-4 p-5 bg-bg-noite border border-bg-bordo rounded-2xl hover:border-ouro-imperial transition-all text-left shadow-sm group"
                  >
                    <div className="p-3 bg-bg-bordo-profundo rounded-xl group-hover:bg-bg-bordo transition-colors">
                      {copySuccess ? <CheckCircle2 className="text-green-500" size={24} /> : <Copy className="text-text-areia" size={24} />}
                    </div>
                    <div>
                      <div className="font-cinzel font-bold text-base text-text-creme">{copySuccess ? 'Copiado!' : 'Copiar Texto'}</div>
                      <div className="text-[10px] text-text-terra font-bold uppercase tracking-widest text-wrap">Área de Transferência</div>
                    </div>
                  </button>

                  <button 
                    onClick={handleDownloadPdf}
                    disabled={isGeneratingPdf}
                    className="w-full flex items-center gap-4 p-5 bg-bg-noite border border-bg-bordo rounded-2xl hover:border-ouro-imperial transition-all text-left shadow-sm group disabled:opacity-60 disabled:cursor-wait"
                  >
                    <div className="p-3 bg-bg-bordo-profundo rounded-xl group-hover:bg-bg-bordo transition-colors">
                      {isGeneratingPdf ? <Loader2 className="text-ouro-imperial animate-spin" size={24} /> : <Download className="text-text-areia" size={24} />}
                    </div>
                    <div>
                      <div className="font-cinzel font-bold text-base text-text-creme">{isGeneratingPdf ? 'Gerando PDF...' : 'Baixar PDF'}</div>
                      <div className="text-[10px] text-text-terra font-bold uppercase tracking-widest text-wrap">{isGeneratingPdf ? 'Convertendo via servidor' : 'Alta Qualidade'}</div>
                    </div>
                  </button>
                </div>

                <div className="pt-8 space-y-4">
                   <p className="text-xs text-text-terra text-center uppercase tracking-widest font-cinzel font-bold">Novos dados?</p>
                   <button 
                    onClick={() => setActiveStep('filler')}
                    className="w-full py-4 px-4 bg-gradient-to-r from-bg-bordo-profundo to-bg-bordo text-text-creme border border-ouro-escuro/30 rounded-2xl font-cinzel font-bold tracking-wider uppercase hover:from-bg-bordo hover:to-bg-bordo-profundo transition-all shadow-xl shadow-bg-ebano hover:-translate-y-0.5"
                  >
                    Reiniciar Campos
                  </button>
                </div>
              </div>

              <div className="lg:col-span-3 space-y-4 order-1 lg:order-2">
                <div className="flex items-center justify-between border-b border-bg-bordo pb-2">
                  <h3 className="text-xs font-cinzel font-bold text-ouro-claro uppercase tracking-[0.3em]">Documento Gerado</h3>
                </div>
                <div 
                  id="document-preview-content" 
                  ref={printRef} 
                  className={`bg-white border border-ouro-escuro/30 rounded-2xl shadow-2xl min-h-[1122px] relative overflow-hidden ql-container ql-snow print:shadow-none print:border-none print:p-0`}
                  style={{
                    backgroundImage: 'url(/a.png)',
                    backgroundSize: '100% 1122px', // A4 page height equivalent
                    backgroundRepeat: 'repeat-y',
                    backgroundPosition: 'top center'
                  }}
                >
                  <div className="relative z-10 h-full flex flex-col" style={{
                    paddingTop: '280px',     // Clear the letterhead header area with extra breathing room
                    paddingBottom: '160px',  // Clear the letterhead footer area with extra breathing room
                    paddingLeft: '28mm',     // Clear the border frame left
                    paddingRight: '28mm',    // Clear the border frame right
                  }}>
                    <div 
                      className={`ql-editor font-serif text-black leading-relaxed text-sm lg:text-[17px] selection:bg-bg-bordo/20 print:p-0 flex-1`}
                      dangerouslySetInnerHTML={{ __html: filledText }}
                    />
                  </div>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <footer className="max-w-7xl mx-auto px-4 py-16 border-t border-bg-bordo mt-20 flex flex-col md:flex-row justify-between items-center gap-6 text-text-terra text-sm relative z-10">
        <div className="flex items-center gap-3">
          <Settings size={14} className="animate-spin-slow text-ouro-imperial" />
          <span className="font-cinzel tracking-wider uppercase text-xs">AutoTermos • Simples Contábil</span>
        </div>
        <div className="flex items-center gap-8 font-cinzel tracking-wider uppercase text-xs">
          <a href="#" className="hover:text-ouro-imperial transition-colors">Segurança</a>
          <a href="#" className="hover:text-ouro-imperial transition-colors">Suporte</a>
          <p>© 2026 Simples Contábil</p>
        </div>
      </footer>
    </div>
  );
}
