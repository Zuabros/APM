using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Aposent_o_matic
{
    public partial class Form1 : Form
    {

	// ---------------------------------
	// VERIFICA SE NÃO EXPIROU
	// ---------------------------------
	void checkexpire()
	{
	 // Developer's shortcut - se Shift estiver pressionado, pula a verificação
	 if (Control.ModifierKeys == Keys.Shift)
		return;

	 DateTime now = DateTime.Now;
	 var expire = new DateTime(2026, 12, 31); // expire date 
	 if (DateTime.Compare(now, expire) > 0) Environment.Exit(1); // exit program
	 if (DateTime.Compare(dbfim.Value, expire) > 0) Environment.Exit(1); // exit program
	}
	public Form1()
        {
            InitializeComponent();
            checkexpire();
            // Incializa valores padrão 
            cbbunit.SelectedIndex = 0; // Define valor padrão para unidade de medida dBA
            cbbcaatividade.SelectedIndex = 5;
            cbbcatrab.SelectedIndex = 0;


            // Inicializa combo bomx do calor com IBUTG
            cbbcaun.SelectedIndex = 1;

            

            // ---------------------------------
            // INICIALIZA LISTBOXES DO QUIMICO
            // ---------------------------------
            // Anexo 53.831/64
            string qn64 = "1.2.1 - arsênico,1.2.2 - berílio,1.2.2 - glicinio,1.2.3 - cádmio,1.2.4 - chumbo,1.2.5 - cromo,1.2.6 - fósforo,1.2.7 - " +
                "manganês,1.2.8 - mercúrio,1.2.9 - Tóxicos inorgânicos da lista da OIT,1.2.10 - poeira de sílica," +
                "1.2.10 - poeira de carvão mineral,1.2.10 - poeira de cimento,1.2.10 - poeira de asbestos (amianto),1.2.10 - poeira de talco," +
                "1.2.11 - hidrocarbonetos tóxicos (ano/eno/ino)-OIT," +
                "1.2.11 - ácidos carboxílicos tóxicos (óico)-OIT,1.2.11 - álcoois tóxicos (ol)-OIT,1.2.11 - cetonas tóxicas (ona)-OIT," +
                "1.2.11 - ésteres tóxicos (oxissais em ato-ila)-OIT,1.2.11 - éteres tóxicos (óxidos oxi)-OIT," +
                "1.2.11 - amidas-amidos tóxicas-OIT,1.2.11 - aminas tóxicas-OIT," +
                "1.2.11 - nitrila e isonitrila (nitrilas e carbolaminas) tóxicas-OIT,1.2.11 - compostos organo-metálicos halogenados-OIT," +
                "1.2.11 - compostos metaloidicos halogenados-OIT,1.2.11 - compostos metalódicos e nitrados-OIT,1.2.11 - cloreto de metila," +
                "1.2.11 - tetracloreto de carbono,1.2.11 - tricloroetileno,1.2.11 - clorofórmio,1.2.11 - bromureto de netila (brometo de metila)," +
                "1.2.11 - bromometano (bromureto de metila)," +
                "1.2.11 - nitrobenzeno,1.2.11 - gasolina,1.2.11 - acetona,1.2.11 - acetatos (metil/etil/propil/butil/amil),1.2.11 - pentano,1.2.11 - metano,1.2.11 - haxano," +
                "1.2.11 - sulfureto de carbono,"+
                // ABAIXO VEM DO DECRETO 83080/79
                "1.2.11 - tintas tóxicas e solventes (pintura a pistola),"+
                "1.2.9 - ouro,1.2.10 - benzeno (benzol),1.2.10 - tolueno (toluol),1.2.10 - xileno (xilol),1.2.10 - bromofórmio,1.2.10 - tricloroetileno," +
                "1.2.10 - tetracloreto de carbono,1.2.10 - carbonilida,1.2.10 - gás de iluminação,1.2.10 - ácido carbônico,1.2.11 - flúor e ácido fluorídrico," +
                "1.2.11 - cloro e ácido clorídrico,1.2.11 - bromo e ácido bromídrico,1.2.11 - revestimentos de niquelagem/cromagem/douração,1.2.11 - " +
                "revestimentos de anodização de alumínio,1.2.11 - monóxido de carbono+metano+gás sulfídrico-esgoto," +
                "1.2.11 - fumos metálicos (solda elétrica e oxiacetileno),1.2.11 - alvejantes e tintura (aplicada à mão),1.2.12 - cimento,1.2.12 - fibrocimento," +
                "1.2.12 - sílica (jateamento de areia),1.2.12 - amianto,1.2.11 - Trabalho em galerias e tanques de esgoto,1.2.10 - Tíner (Thinner®: tolueno + etilacetato + etanol)";
            // LISTA ABAIXO RETIRADA DA LISTA DA (RTS da OIT)
            string qRTS = "acetaldeido,ácido acético,anidrido acético,acetona,acroleina,acrilonitrato,alilalcool (prop-2-en-1-ol),amilpropildissulfeto," +
                "amonia,amilacetato,amilalcool (isoamilalcool),anilina,arsina,benzeno (benzol/benzina),cloreto de benzil,butadieno,butanona (metiletilcetona)," +
                "butilacetato (n-butilacetato),butilalcool (n-butanol),butilamina,butil cellosolve (2-butoxietanol),dióxido de carbono,dissulfeto de carbono," +
                "monóxido de carbono,tetracloreto de carbono (tetraclorometano),etoxietanol (cellosolve/2-etoxietanol),acetato de 2-etoxietanol,gás cloro (Cl2)," +
                "trifluoreto de cloro,clorobenzeno (monoclorobenzeno),clorofórmio (triclorometano),1-cloro-1-nitropropano,cloropreno (2-cloro-1-3-butadieno)," +
                "cresol (metilfenol/hidroxitoluneno e todos os isômeros),ciclohexano,cicloexanol,ciclohexanona,ciclopropano,alcool diacetona (4-hidroxi-4-" +
                "metil-2-pentanona),diborano,o-diclorobenzeno (orto-diclorobenzeno),diclorofluormetano,1-1-dicloroetano,1-2-dicloroetileno,dicloroetil eter,dicloromonofluormetano," +
                "1-1-dicloro-1-nitroetano,diclorotetrafluoretano,dietilamina,difluorodibromometano,diisobutil cetona (4-heptanona 2-6-dimetil),dimetilanilina " +
                "(n-dimetilanilina),dimetilsulfato,dioxano (dióxido de dietileno),etilacetato (acetato de etila),etil alcool,etilamina,etilbenzeno,etilbrometo,etilcloreto," +
                "etiléter,etil formato,etil silicato,etil clorohidrina,etilenodiamina,etileno dibrometo,etileno dicloreto,etileno imina,óxido de etileno," +
                "gás flúor (F2),fluortriclorometano,formaldeído (formol),gasolina,heptano (n-heptano),hexano,hexanona (metilbutilcetona),hexona (metil" +
                "isobutilcetona),hidrazina,ácido bromídrico (HBr),ácido clorídrico (HCl),ácido cianídico (HCN),ácido fluorídrico (HF),peróxido de hidrogênio 90% " +
                "(água oxigenada),ácido sulfídrico (H2S),seleneto de hidrogênio (H2Se),iodo (I2),ISO,isoforona,isopropilamina,oxido de mesitil,metilacetato (acetato de metila)" +
                ",acetileno de metila (metilacetileno),metil alcool (metanol),brometo de metila,metil cellosolve (2-metoxietanol),acetato de metilcellosolve (" +
                "acetato de etilenoglicol monometil éter),cloreto de metila,dimetoximetano (metilal),metil cloroformio,metilciclohexano,metilciclohexanol," +
                "metilciclohexanona, formato de metila,metilisobutilcarbinol (metilamil alcool),cloro metileno (dihidroclorometano), nafta (alcatrão),nafta (" +
                "pétróleo),niquel carbonil,p-nitroanilina,nitrobenzeno,nitroetano,dioxido de nitrogenio,nitroglicerina,nitrometano,2-nitropropano,nitrotolueno," +
                "octano,ozônio,pentano,pentanona (metilpropil cetona),percloroetileno (tetracloroetileno),fenol,fenohidrazina,fosgenio (COCL2),fosfina (PH3)," +
                "tricloreto de fósforo (PCl3),propilacetato,propil alcool (alcool isopropílico),propil eter (éter isopropílico),dicloropropano (dicloreto de propileno / " +
                "dicloropropileno),propileno imina,piridina,quinona,stibina,solvente de Stoddard,feniletileno, monomero estireno,dióxido de enxofre," +
                "hexafluoreto de enxofre,fluoreto de enxofre VI,monocloreto de enxofre,pentafluoreto de enxofre,p-tributiltolueno,1-1-2-2-tetracloroetano (tetracloro de acetileno)," +
                "tetranitrometano,tolueno (toluol),o-toluidina (ortotoluidina),tricloroetileno,trifluormonobromoetano,turpentina,cloroetileno (cloreto de vinila)," +
                "xileno (xilol),";
            // POEIRAS TÓXICAS, FUMOS E NÉVOAS DA OIT
            qRTS += "aldrina (hexacloro hexahidro metanonaftaleno),amato (sufamato de amonia),antimônio, arsênio,bário (compostos solúveis), fumos " +
                "de óxido de cádmio, clordano (octacloro tetrahidro metanoindeno),óxido difenil clorado,clorodifenil (42% cloro),ácido crômico e cromatos," +
                "herbicida Crag (sesone),cianeto (CN),ácido 2-4-diclorofenoxiacético,dieldrina (hexacloro epoxi octahidro dimetanonaftaleno),bromureto de netila," +
                "dinitrotolueno," +
                "dinitro-o-cresol (dinitroortocresol),EPN (o-etil-o-p-nitrofenil tionobenzenofosfonato),poeira de ferro/vanádio,hidroquinona,fumos de óxido " +
                "de ferro,chumbo,lindano (hexaclorociclohexano;isomero gama),fumos de óxido de magnésio,Malathion (dimetil difosfonato de dietil mercaptocuccinato)," +
                "manganês,mercúrio (compostos orgânicos),mercúrio (inorgânico),metoxicloro (metoxifenil tricloroetano),molibdênio (componentes solúvel e insolúvel)," +
                "Parathion (dietil p-nitrofenil tiofosfato),pentacloronaftaleno,pentaclorofenol,fósforo (amarelo),pentacloreto de fósforo, pentassulfeto de fósforo," +
                "ácido pícrico,compostos deselênio,hidróxido de sódio,ácido sulfúrico,TEDP (tetraetil ditionopirofosfato),TEPP (tetraetil pirofosfato),tellurio," +
                "tetril (trinitrofenilmetilnitramina),dióxido de titânio,tricloronaftaleno,trinitrotolueno (TNT),urânio (solúvel e insolúvel),poeiras e fumos de vanádio," +
                "fumos de óxido de zinco,compostos de zircônio (como Zr),poeira de óxido de alumínio,poeira de asbestos (amianto),poeira de mica" +
                " (abaixo de 5% de sílica livre),poeira de cimento Portland,poeira de talco,poeira de carborundum,poeira de ardósia,poeira de pedra sabão";
            // LINACH GRUPO 1
            string qlinachhard = "1.0.1 - arsênio/arsênico e compostos inorgânicos,1.0.2 - asbestos ou amianto em todas as formas,1.0.3 - benzeno/benzidina/" +
                    "benzopireno,1.0.4 - berílio e seus compostos,1.0.6 - cádmio e compostos de cádimio,1.0.7 - breu de alcatrão de hulha," +
                    "1.0.9 - bifenil policlorado (Ascarel),1.0.9 - cloreto de vinila,1.0.9 - 2-3-7-8 tetraclorodibenzeno-para-dioxina,1.0.9 - tricloroetileno," +
                    "1.0.10 - compostos de cromo (VI),1.0.12 - fósforo 32 (como fosfato),1.0.17 - óleos de xisto,1.0.18 - poeiras de sílica (quartzo/cristalina/cristobalita)," +
                    "1.0.19 - 4-aminobifenila,1.0.19 - azatioprina," +
                    "1.0.19 - ciclofosfamida,1.0.19 - clorambucil,1.0.19 - dietilestilbestrol,1.0.19 - benzopireno (benzapireno),1.0.19 - éter bisclorometílico (bisclorometiléter)," +
                    "1.0.19 - éter metílico de clorometila (clorometiléter),1.0.19 - fenacetina (mistura de analgésico contendo),1.0.19 - orto-toluidina,1.0.19 - 1-3 butadieno," +
                    "1.0.19 - óxido de etileno,1.0.19 - benzidina,1.0.19 - 2-naftilamina";
            // DECRETO 3.048/99 ANEXO IV - QUALITATIVOS
            string qanexoivqual = "1.0.1 - arsênio e seus compostos tóxicos,1.0.3 - benzeno e seus compostos tóxicos," +
                    "1.0.4 - berílio e seus compostos tóxicos,1.0.6 - cádmio e seus compostos tóxicos,1.0.7 - carvão mineral e seus derivados,1.0.7 - piche (derivado do carvão mineral)," +
                    "1.0.7 - alcatrão (derivado do carvão mineral),1.0.7 - betume (derivado do carvão mineral),1.0.7 - breu (derivado do carvão mineral),1.0.7 - antraceno (derivado do carvão mineral)," +
                    "1.0.7 - parafina (derivado do carvão mineral),1.0.7 - coque (derivado do carvão mineral)," +
                    "1.0.10 - cromo e seus compostos tóxicos,1.0.12 - fósforo e seus compostos tóxicos," +
                    "1.0.13 - iodo (fabricação e emprego industrial),1.0.15 - mercúrio orgânico (metil e etilmercúrio)," +
                    "1.0.17 - petróleo/xisto/gás natural e derivados,1.0.17 - gasolina/óleo diesel (derivado do petróleo),1.0.17 - querosene (derivado do petróleo),1.0.19 - bisclorometil," +
                    "1.0.16 - Níquel (níquel carbonila e sulfeto de níquel)," +
                    "1.0.19 - mercaptanos exceto n-butilmercaptano/butanotiol/etilmercaptano/etanotiol,1.0.19 - aminas aromáticas,1.0.19 - auramina,1.0.19 - bisclorometileter," +
                    "1.0.19 - biscloroetileter,1.0.19 - clorometileter,1.0.19 - nitronaftilamina,1.0.19 - benzopireno,1.0.19 - creosoto,1.0.19 - 4-aminodifenil," +
                    "1.0.19 - betanaftilamina,1.0.19 - aminobifenila,1.0.19 - 4-dimetilaminazobenzeno,1.0.19 - betapropilactona,1.0.19 - dianizidina," +
                    "1.0.19 - diclorobenzidina,1.0.19 - metileno-ortocloroanilina (MOCA),1.0.19 - nitrosamina,1.0.19 - ortotoluidina,1.0.19 - propanosultona," +
                    "1.0.3 - etilbenzeno,1.0.19 - 1-cloro-2-4-nitrodifenil,1.0.19 - 1-cloro-2-3-epoxipropano (epicloridrina),1.0.19 - azatioprina,1.0.19 - clorambucil," +
                    "1.0.19 - dietilestilbestrol,1.0.19 - dimetilsulfato,1.0.19 - etilenotiuréia,1.0.19 - fenacetina,1.0.19 - iodeto de metila,1.0.19 - etilnitrosuréias," +
                    "1.0.19 - oximetalona,1.0.19 - procarbazina,1.0.19 - óxido de etileno,1.0.19 - demetanosulfonato (Busulfan/Mileran),1.0.19 - dietil-bestrol"+
                    "1.0.19 - 4-aminobifenila,1.0.8 - chumbo e seus compostos tóxicos," +
                    "1.0.19 - benzopireno (benzapireno),1.0.19 - ciclofosfamida," +
                    "1.0.19 - benzidina,1.0.19 - n-hexano,1.0.19 - dietilsulfato";
            string qivquant = 
                "1.0.2 - asbestos crisotila," + // 0
                "1.0.5 - brometo de etila (bromoetano)," + //1 156 695
                "1.0.5 - brometo de metila (bromometano)," + //2 12 47
                "1.0.5 - bromo," + //3 0,08 0,6
                "1.0.5 - bromofórmio," + //4 0,4 4
                "1.0.5 - dibromometano," + //5 16 110
                "1.0.9 - ácido clorídrico," + //6 4 5,5
                "1.0.8 - chumbo e seus compostos tóxicos," + //7 - 0,1
                "1.0.9 - cloreto de etila (cloroetano)," + //8 780 2030
                "1.0.9 - cloreto de metila," + //9 78 165
                "1.0.9 - cloreto de metileno," + //10 156 560
                "1.0.9 - cloreto de vinila (cloroetílico)," + //11 156 398
                "1.0.9 - cloreto de vinilideno," + //12 8 31
                "1.0.9 - cloro," + //13 0,8 2,3
                "1.0.9 - cloro 1-nitropropano," + //14 16 78
                "1.0.9 - clorobenzeno (cloreto de fenila)," + //15 
                "1.0.9 - clorobromometano," + //16  156 820
                "1.0.9 - clorodifluometano (freon 22)," + //17 780 2730
                "1.0.9 - clorofórmio," + //18 20 94
                "1.0.19 - cloroprene (cloropreno)," + //19 20 70
                "1.0.9 - 1-1-dicloro-1-nitroetano," + //20 8 47
                "1.0.9 - percloroetileno," + //21 78 525
                "1.0.11 - dissulfeto de carbono," + //22 16 47
                "1.0.14 - manganês e seus compostos tóxicos," + //23 0 - 5mg/m3 extração e 1mg/m3 atividade industrial 
                "1.0.16 - níquel carbonila," + //24 0,04 0,25
                "1.0.7 - negro de fumo," + //25 - 3,5
                "1.0.18 - sílica livre," + // 26
                "1.0.19 - estireno," + //27 78 328
                "1.0.19 - 1-3-butadieno," + //28 780 1720
                "1.0.19 - acrilonitrila (acronitrila)," + //30 16 35
                "1.0.19 - diisocianato de tolueno (TDI/DTI)," + //31 0,016 0,11
                "1.0.19 - etilenoamina," + //32 0,4 0,8
                "1.0.9 - fosgênio (fluoreto de carbonila)," +//33 0,08 0,3
                "1.0.19 - n-butilmercaptano (butanotiol)," + //34  0,4 1,2
                "1.0.19 - etanotiol (etil mercaptano)," + //35 0,4 0,8
                "1.0.19 - 1-4-butanodiol (dihidroxibutano)," + //38 nao tem nr 15
                "1.0.19 - tetrametilenoglicol (butanodiol)"; //39 nao tem na nr 15


            string limites =
            "2,0 f/cm³[" + // asbestos
            "156ppm / 695mg/m³[" + // bromoetano
            "12ppm / 47mg/m³[" + // bromometano
            "0,08ppm / 0,6mg/m³[" + // bromo
            "0,4ppm / 4mg/m³[" + // bromoformio
            "16ppm / 110mg/m³[" + // dibromometano 
            "4ppm / 5,5mg/m³[" + // acido cloridrico 
            "7ppm / 0,1mg/m³[" + // chumbo 
            "780ppm / 2030mg/m³[" + // cloreto de etila
            "78ppm / 165mg/m³[" + // cloreto de metila 
            "156ppm / 560mg/m³[" + // cloreto de metileno 
            "156ppm / 398mg/m³[" + // cloreto de vinila 
            "8ppm / 31mg/m³[" + // cloreto de vinilideno 
            "0,8ppm / 2,3mg/m³[" + // cloro 
            "16ppm / 78mg/m³[" + // cloro 1 nitropropano 
            "59ppm / 275mg/m³[" + // cloreto de fenila
            "156ppm / 820mg/m³[" + // clorobromometano 
            "780ppm / 2730mg/m³[" + // clorodifluometano 
            "20ppm / 94mg/m³[" + // cloroformio 
            "20ppm / 70mg/m³[" + // cloropreno
            "8ppm / 47mg/m³[" + // dicloro nitroetano
            "78ppm / 525mg/m³[" + // percloroetileno  
            "16ppm / 47mg/m³[" + // dissulfeto de carbono 
            "5mg/m³ extração e 1mg/m³ indústria[" + // manganês 
            "0,04ppm / 0,25mg/m³[" + // niquel carbonila
            "25ppm / 3,5mg/m³[" + // negro de fumo 
            "anexo 12 da NR-15[" + // silica livre
            "78ppm / 328mg/m³[" + // estirno
            "780ppm / 1720mg/m³[" + // 1-3 butadieno 
            "16ppm / 35mg/m³[" + // ACRILONITRILA / acronitrila
            "0,016ppm / 0,11mg/m³[" + // DTI
            "0,4ppm / 0,8mg/m³[" + // ETILENOAMINA
            "0,08ppm / 0,3mg/m³[" + // FOSGENIO - FLUORETO DE CARBONILA
            "0,4ppm / 1,2mg/m³[" + // butanotiol
            "0,4ppm / 0,8mg/m³[" + // etilmercaptano 
            "?[" + // butanodiol, tetrametilenoglicol, dihidroxibutano 
            "?"; // butanodiol ? 

            //string CID = 
            lim = s2dic(qivquant, limites);

            qoit = s2arr(qRTS); // read from hard-coded string to array 
            arr2lb(qoit, ref lbqoit); // passa para a listbox 
            qa64 = s2arr(qn64); // read from hard-coded string to array 
            arr2lb(qa64, ref lbq64); // passa para a listbox 
            qlinach = s2arr(qlinachhard);
            arr2lb(qlinach, ref lbqlinach); // passa para a listbox 
            qaIVqual = s2arr(qanexoivqual);
            arr2lb(qaIVqual, ref lbqanexoivqual); // passa para a listbox 
            qaivquant = s2arr(qivquant);
            arr2lb(qaivquant, ref lbqivquant); // passa para a listbox 
            //---------------------------------------------
            // QUIMICOS - ADICIONA LIMITES AO DICIONÁRIO
            //---------------------------------------------            
            
            //---------------------------------
            // Converte array de strings em listbox 
            void arr2lb(string[] arr,ref ListBox lb)
            {
                foreach (var item in arr)
                    lb.Items.Add(item);
            }

            //---------------------------------
            // Gera dictionary de 2 strings. 
            Dictionary<string, string> s2dic(string idx,string words)
            {
                string[] tidx = idx.Split(',');
                for (int i = 0; i < tidx.Length; i++)
                    tidx[i] = tidx[i].Trim();
                string[] twords = words.Split('[');
                for (int i = 0; i < twords.Length; i++)
                    twords[i] = twords[i].Trim();
                Dictionary<string, string> tdic = new Dictionary<string, string>();
                if (tidx.Length != twords.Length) 
                    {   
                    fala($"Erro na função s2dic: array size mismatch TIDX={tidx.Length} TWORDS={twords.Length}");
                    return tdic; }
                for (int i=0;i<tidx.Length;i++)
                    tdic.Add(tidx[i], twords[i]);
                return tdic;
            }
            //---------------------------------
            // Converte string para array de strings
            string[] s2arr(string s)
            {
                string[] tarr = s.Split(',');
                for (int i = 0; i < tarr.Length; i++)
                    tarr[i] = tarr[i].Trim();
                Array.Sort(tarr);
                return tarr;
            }
        }
        //---------------------------------
        // Pinta os periodos
        //---------------------------------
        void pinta(int idx, int qual = 0)
        {
            var lab0 = labq1;
            var lab1 = labq2;
            var lab2 = labq3;
            var lab3 = labq4;
            var lab4 = labq5;
            
            if (qual == 0)
            {
                lab0 = labq1;
                lab1 = labq2;
                lab2 = labq3;
                lab3 = labq4;
                lab4 = labq5;
            }
            else
            {
                lab0 = labq21;
                lab1 = labq22;
                lab2 = labq23;
                lab3 = labq24;
                lab4 = labq25;
            }
            
            
            // COLORE PRIMEIRA DATA
            if (DateTime.Compare(dbinicio.Value, j4.fim) > 0) // começa depois 
                lab0.ForeColor = Color.Gray;
            else
            {
                if (qagflag[idx,6] || qagflag[idx, 7]) lab0.ForeColor = Color.Red; // 6=eventual 7=generico 
                else if (qagflag[idx, 0] || qagflag[idx, 1]) lab0.ForeColor = Color.Green; // 0=decreto64  1=OIT
                else lab0.ForeColor = Color.Red; // 6=eventual 7=generico 
            }
            // COLORE SEGUNDA
            if (DateTime.Compare(dbfim.Value, j4.inicio) < 0 // termina antes 
                || (DateTime.Compare(dbinicio.Value, new DateTime(2003, 11, 18)) > 0)) // começa depois 
                lab1.ForeColor = Color.Gray;
            else
            {
                if (qagflag[idx, 6] || qagflag[idx, 7]) lab1.ForeColor = Color.Red; // 6=eventual 7=generico 
                else if (qagflag[idx, 2]) lab1.ForeColor = Color.Green; // 97 quali   
                else if (qagflag[idx, 3] && qagflag[idx, 4]) lab1.ForeColor = Color.Green; // 3=quant 4=quant ok
                else lab1.ForeColor = Color.Red; // 6=eventual 7=generico 
            }
            // COLORE TERCEIRA
            if (DateTime.Compare(dbfim.Value, new DateTime(2003, 11, 19)) < 0 // termina antes 
                || (DateTime.Compare(dbinicio.Value, new DateTime(2003, 12, 31)) > 0)) // começa depois 
                lab2.ForeColor = Color.Gray;
            else
            {
                if (qagflag[idx, 6] || qagflag[idx, 7]) lab2.ForeColor = Color.Red; // 6=eventual 7=generico 
                else if (qagflag[idx, 2]) lab2.ForeColor = Color.Green; // 97 quali
                else if (qagflag[idx, 3] && qagflag[idx, 4]) lab2.ForeColor = Color.Green; // 3=quant 4=quant ok
                else lab2.ForeColor = Color.Red; // 6=eventual 7=generico 
            }
            // VERIFICA PENULTIMA DATA
            if (DateTime.Compare(dbfim.Value, new DateTime(2004, 10, 08)) < 0 // termina antes 
                | (DateTime.Compare(dbinicio.Value, new DateTime(2004, 10, 07)) > 0)) // começa depois
                lab3.ForeColor = Color.Gray;
            else
            {
                if (qagflag[idx, 6] || qagflag[idx, 7]) lab3.ForeColor = Color.Red; // 6=eventual 7=generico 
                else if (qagflag[idx, 2]) lab3.ForeColor = Color.Green; // 97 quali
                else if (qagflag[idx, 3] && qagflag[idx, 4]) lab3.ForeColor = Color.Green; // 3=quant 4=quant ok
                else lab3.ForeColor = Color.Red; // 6=eventual 7=generico 
            }
            // VERIFICA ULTIMA DATA
            if (DateTime.Compare(dbfim.Value, new DateTime(2014, 10, 08)) < 0) // termina antes 
                lab4.ForeColor = Color.Gray;
            else
            {
                if (qagflag[idx, 6] || qagflag[idx, 7]) lab4.ForeColor = Color.Red; // 6=eventual 7=generico 
                else if (qagflag[idx, 2]) lab4.ForeColor = Color.Green; // 97+ quali
                else if (qagflag[idx, 5]) lab4.ForeColor = Color.Green; // LINACH
                else if (qagflag[idx, 3] && qagflag[idx, 4]) lab4.ForeColor = Color.Green; // 3=quant + 4=quant ok
                else lab4.ForeColor = Color.Red; // 6=eventual 7=generico 
            }
        }
        string[] qag = new string[20]; // nome dos agentes 
        bool[,] qagflag = new bool[30,20]; // flags do agente [agente index,flag index]
        // 0==decreto 64 1==oit  2==97quali   3==97quant  4==97quantOK  5==LINACH  6==eventual  7==generico
        // 8==aceitar metodologia   9==NR15   10==NHO

        string[] qoit; // lista da OIT pré 97 
        string[] qa64; // lista do INSS 1964
        string[] qlinach; // linach (cancerigenos, qualitativo)
        string[] qaIVqual; // Decreti 3048/99 anexo IV qualitativo 
        string[] qaivquant; // anexo IV quantitativo
        Dictionary<string, string> lim;
        private void Form1_Load(object sender, EventArgs e)
        {


        }
        // ------------------------------
        // CONSTANTES
        // ------------------------------
        const string ENTER = "\r\n";
        const string ENTERS = "\r\n\r\n";
        const string SP = " ";
        const string PT = ".";
        const string VG = ",";
        const string NAOPERMA = "A exposição ao agente foi ocasional/intermitente, conforme descritivo da profissiografia da" +
            " atividade laboral (campo 14.2 do PPP). ";
        DateTime PUBNHO = new DateTime(2001, 1, 1);

        // ------------------------------
        // ESTRUTURAS DE DADOS (STRUCTS) 
        // ------------------------------

        struct janela          // Estrutura dos períodos do quadro 14 do manual 
        {
            public int num;             // numero 
            public DateTime inicio;     // Data inicio do periodo
            public DateTime fim;        // Data fim do periodo
            public double limite;     // Valor limite da tolerância acústica
            public int metodo;             // 0 = NR-15 anexo I  2 = NHO FUNDACENTRO  12 = NR-15 ou NHO 01 FUNDACENTRO  99 = outro (inválido) 
            public string lei;             // Legislação aplicável ao período 
            public bool EPC;               // TRUE = precisa EPC; FALSE= não precisa
            public bool EPI;               // TRUE = precisa EPI; FALSE= não precisa 
            public List<string> codigo;          // Códigos utilizados pelo INSS para o período 
            public List<string> codigo2;         // Códigos alternativos especificos de algumas funções
            public bool outros;
            public bool antigo;

        }
        struct laudo // estrutura com o laudo para a janela especifica 
        {
            public DateTime inicio; // data de inicio da cobertura do laudo 
            public DateTime fim; // data termino da cobertura do laudo 
        }
        struct enquadra // estrutura com o laudo para a janela especifica 
        {
            public bool metodo; // metodo está ok 
            public bool unidade; // unidade está ok 
            public bool limite; // acima do limite 
            public bool permanencia; // permanente 
            public bool tudo; // enquadrou tudo
            public bool epoca; // epoca de enquadramento, para eletreicidade, frio e umidade antes de 05/03/97
            public bool outros;
            public bool especial; // flag para condições espciais, como NEN
        }

        // ---------------------------------------
        // VARIÁVEIS hard-coded (dados tabulados)
        // ---------------------------------------

        List<bool> passou = new List<bool> { false, false, false, false, false, false, false }; // lista de periodos enquadrados ou não 

        janela j2 = new janela  // Periodo até 13/10/1996, engloba 2 periodos iniciais similares 
        {
            num = 0,
            inicio = new DateTime(1964, 01, 01),
            fim = new DateTime(1996, 10, 13),
            limite = 80,
            metodo = 1,
            lei = "Decreto nº 53.831, de 1964",
            EPI = false,
            EPC = false,
            // 0=ruído, 1=eletriciade, 2=químico, 3=biologico, 4=calor, 5=rni, 6=frio, 7=umidade, 8=vibração
            codigo = new List<string> { "1.1.6", "1.1.8", "", "1.3.2", }, // 132=biologicos humanos
            codigo2= new List<string> { "", "", "", "1.3.1", }, // biologicos veterinario 
            outros = true,
            antigo = true,
        };
        janela j3 = new janela  // Periodo 14/10/96 a 05/03/97
        {
            num = 1,
            inicio = new DateTime(1996, 10, 14),
            fim = new DateTime(1997, 03, 05),
            limite = 80,
            metodo = 1,
            lei = "Decreto nº 53.831, de 1964 e MP nº 1.523, de 1996",
            EPI = false,
            EPC = true,
            // 0=ruído, 1=eletriciade, 2=químico, 3=biologico, 4=calor, 5=rni, 6=frio, 7=umidade, 8=vibração
            codigo = new List<string> { "1.1.6", "1.1.8", "", "1.3.2", },
            codigo2 = new List<string> { "", "", "", "1.3.1", }, // biologicos veterinario 
            outros = true,
            antigo = true,

        };
            janela j4 = new janela  // Periodo 06/03/97 a 02/12/98
            {
            num = 2,
            inicio = new DateTime(1997, 03, 06),
            fim = new DateTime(1998, 12, 02),
            limite = 90,
            metodo = 1,
            lei = "Decreto nº 2.172, de 1997",
            EPC = true,
            EPI = false,
                                    // 0=ruído, 1=eletriciade, 2=químico, 3=biologico, 4=calor, 5=rni, 6=frio, 7=umidade, 8=vibração
            codigo = new List<string> { "2.0.1", "1.1.8", "", "3.0.1", },
            codigo2 = new List<string> { "", "", "", "3.0.1", }, // alternativos 
            outros = false,
            antigo = false,

        };
        janela j5 = new janela  // Periodo 03/12/98 a 6/5/99
        {
            num = 3,
            inicio = new DateTime(1998, 12, 03),
            fim = new DateTime(1999, 5, 6),
            limite = 90,
            metodo = 1,
            lei = "Decreto nº 2.172, de 1997 e Lei nº 9.528, de 1997",
            EPC = true,
            EPI = true,
                                    // 0=ruído, 1=eletriciade, 2=químico, 3=biologico, 4=calor, 5=rni, 6=frio, 7=umidade, 8=vibração
            codigo = new List<string> { "2.0.1", "1.1.8", "", "3.0.1", },
            codigo2 = new List<string> { "", "", "", "3.0.1", }, // alternativos 
            outros = false,
            antigo = false,

        };
        janela j6 = new janela  // Periodo 06/03/97 a 02/12/98
        {
            num = 4,
            inicio = new DateTime(1999, 5, 7),
            fim = new DateTime(2003, 11, 18),
            limite = 90,
            metodo = 1,
            lei = "Decreto nº 3.048, de 1999",
            EPC = true,
            EPI = true,
                                    // 0=ruído, 1=eletriciade, 2=químico, 3=biologico, 4=calor, 5=rni, 6=frio, 7=umidade, 8=vibração
            codigo = new List<string> { "2.0.1", "1.1.8", "", "3.0.1", },
            codigo2 = new List<string> { "", "", "", "3.0.1", }, // alternativos 
            outros = false,
            antigo = false,
        };
        janela j7 = new janela  // Periodo 19/11/03 a 31/12/03
        {
            num = 5,
            inicio = new DateTime(2003, 11, 19),
            fim = new DateTime(2003, 12, 31),
            limite = 85,
            metodo = 12,
            lei = "Decreto nº 3.048, de 1999, modificado pelo Decreto nº 4.882, de 2003",
            EPC = true,
            EPI = true,
                                     // 0=ruído, 1=eletriciade, 2=químico, 3=biologico, 4=calor, 5=rni, 6=frio, 7=umidade, 8=vibração
            codigo = new List<string> { "2.0.1", "1.1.8", "", "3.0.1", },
            codigo2 = new List<string> { "", "", "", "3.0.1", }, // alternativos 
            outros = false,
            antigo = false,
        };
        janela j8 = new janela  // Periodo 01/01/04 - atual 
        {
            num = 6,
            inicio = new DateTime(2004, 1, 1),
            fim = new DateTime(2030, 01, 01),
            limite = 85,
            metodo = 2,
            lei = "Decreto nº 3.048, de 1999, modificado pelo Decreto nº 4.882, de 2003 e IN 99/INSS/DC, de 2003",
            EPC = true,
            EPI = true,
                                     // 0=ruído, 1=eletriciade, 2=químico, 3=biologico, 4=calor, 5=rni, 6=frio, 7=umidade, 8=vibração
            codigo = new List<string> { "2.0.1",     "1.1.8",       "",      "3.0.1",      },
            codigo2 = new List<string> { "", "", "", "3.0.1", }, // alternativos 
            outros = false,
            antigo = false,

        };

        janela j8b = new janela  // Periodo 01/01/04 - 2014
        {
            num = 6,
            inicio = new DateTime(2004, 1, 1),
            fim = new DateTime(2014, 10, 07),
            limite = 85,
            metodo = 2,
            lei = "Decreto nº 3.048, de 1999, modificado pelo Decreto nº 4.882, de 2003 e IN 99/INSS/DC, de 2003",
            EPC = true,
            EPI = true,
            // 0=ruído, 1=eletriciade, 2=químico, 3=biologico, 4=calor, 5=rni, 6=frio, 7=umidade, 8=vibração
            codigo = new List<string> { "2.0.1", "1.1.8", "", "3.0.1", },
            codigo2 = new List<string> { "", "", "", "3.0.1", }, // alternativos 
            outros = false,
            antigo = false,

        };


        janela j9 = new janela  // Periodo 2014+
        {
            num = 7,
            inicio = new DateTime(2014, 10, 8),
            fim = new DateTime(2030, 01, 01),
            limite = 85,
            metodo = 2,
            lei = "Dec. 3.048/99, modificado pelo Dec. 4.882/03 e IN 99/INSS/DC/2003 e Port. MTE/MS/MPS nº9/2014",
            EPC = true,
            EPI = true,
            // 0=ruído, 1=eletriciade, 2=químico, 3=biologico, 4=calor, 5=rni, 6=frio, 7=umidade, 8=vibração
            codigo = new List<string> { "2.0.1", "1.1.8", "", "3.0.1", },
            codigo2 = new List<string> { "", "", "", "3.0.1", }, // alternativos 
            outros = false,
            antigo = false,

        };
        // ------------------------------
        // VARIÁVEIS dinâmicas
        // ------------------------------
        DateTime inicio; // data de inicio do periodo a avaliar 
        DateTime fim; // data final do periodo a avaliar 
        double limite; // limite informado (valor único)
        double min; // limite informado (valor minimo)
        double max; // limite informado (valor máximo) 
        int metodo; // método informado, se padronizado:   // 0 = NR-15 anexo I  2 = NHO FUNDACENTRO  12 = NR-15 ou NHO 01 FUNDACENTRO  99 = outro (inválido) 
        int unidade = 0; // 0 = dB(A), 1 = dB
        enquadra ruido; // flags do ruido
        enquadra ele;
        string unidtxt; // unidade, discursiva;
        double elevolts, elemin, elemax;
        // ------------------------------
        // FLAGS DE ENQUADRAMENTO POR AGENTE
        // ------------------------------
 
            
        // ------------------------------
        // MÉTODOS (funções)  
        // ------------------------------
        // ------------------------------
        // LEDOUBLE -> LÊ DOUBLE DE STRING 
        // ------------------------------
        double ledouble(string campo) // Lê um numero string de um campo, substitui ponto por vírgula e converte para double 
        {
            return Convert.ToDouble(campo.Replace('.', ',')); // necessário para aceitar input com separador decimal ponto ou virgula 
        }
        // ---------------------------------------------------------------------------------------
        // Função vesenquadra - verifica se enquadra, retorna 1 = sim 0 = nao 2 = não pela unidade 
        // ---------------------------------------------------------------------------------------
        void vesenquadra(janela j)
        {
   ruido.permanencia = (cbruidoperma.Checked);
   ruido.limite = (limite > j.limite); // verifica se o limite informado é maior que o limite do período
	 ruido.unidade =  (!cb_unidade_errada.Checked) ;  // unidade correta
	 ruido.metodo =
		 (metodo == j.metodo) // método exato 
		 || ((metodo == 1 || metodo == 2) && j.metodo == 12) // aceita qualquer um dos dois 
		 || ((metodo == 2 || metodo == 12) && j.metodo == 1 && cblayout.Checked); // Aceita método novo na época velha (layout mantido)
	 ruido.especial = cb_usou_NEN.Checked || metodo != 2; // seta true se usou NEN ou não é NHO


   // CASO PARTICULAR DO NEM 
   if (j.num > 5 && metodo == 2 && !cb_usou_NEN.Checked) // NHO 01, mas não usou NEN
   {
    ruido.especial = false;
   }

		ruido.tudo = (
		 ruido.limite &&
		 ruido.metodo &&
		 ruido.unidade &&
		 ruido.permanencia &&
		 ruido.especial // NEN
	 ); // consolida todas as verificações em uma única flag

	}
        
        // ---------------------------------------------------------------------------------------
        // Função testaele -> verifica se enquadra a eletricidade  
        // ---------------------------------------------------------------------------------------
        void testaele(janela j)
        {


            elevolts = ledouble(tbeleunico.Text);
            elemin = ledouble(tbelemin.Text);
            elemax = ledouble(tbelemax.Text);

            if (rbeleminmax.Checked) elevolts = elemin;


            if (!j.outros) { ele.epoca = false; ele.tudo = false; }
            if (!cbeleperma.Checked) { ele.permanencia = false; ele.tudo = false; }
            if (j.outros) // periodo enquadravel, avalia limite 
            {
                if (rbelequalitativo.Checked) { ele.metodo = false; ele.tudo = false; }
                else if (rbeleacima250.Checked && cbelenaoacima250.Checked) { ele.limite = false; ele.tudo = false; }
                else if (rbeleacima250.Checked && !cbelenaoacima250.Checked) { elevolts = 251; } // "acima de 250..."
                else if (elevolts <= 250) { ele.limite = false; ele.tudo = false; }
                else if (cbeleoutro.Checked) { ele.outros = false; ele.tudo = false; }
                else if (rbeleunidade.Checked) { ele.unidade = false; ele.tudo = false; }
            }

        }
        // -----------------------------------------------------------------------------
        // Função SIMPLE - tira case e outras coisas de uma string 
        // -----------------------------------------------------------------------------
        string simple (string s)
        {
            s = s.ToLower();
            string rep = "çáàâãéêíôóõú-()";
            string sub = "caaaaeeiooou   ";
            string sf = "";
            for (int i = 0; i < s.Length; i++)
            {
                bool achou = false;
                for (int j = 0; j < rep.Length; j++)
                    if (s[i] == rep[j]) { sf += sub[j]; achou = true;  break; }
                if (!achou) sf += s[i];
            }
            while (sf.Contains(' ')) sf = sf.Replace(" ", "");
            return sf;
        }

        // -----------------------------------------------------------------------------
        // Função FILTRA - aplica filtro em uma listbox
        // -----------------------------------------------------------------------------
        void filtra(string s, string[] arr,ref ListBox lb, bool ampla=true)
        {
            if (ampla) s = simple(s);
            lb.Items.Clear();
            foreach (var chemical in arr)
                {
                string tmp = chemical;
                if (ampla) tmp = simple(tmp);
                if (tmp.Contains(s)) lb.Items.Add(chemical);
                }
           
        }
        // -----------------------------------------------------------------------------
        // Função transcreve_metodo - descodifica e retorna a string do método de aferição do som 
        // -----------------------------------------------------------------------------
        string getmetodo(int method, bool literal = false)
        {
            
            //if (rbmetodo2.Checked) return tbmetodo.Text;
            if (method == 1) return "NR-15 do MTE";
            else if (method == 12)
            {
                if (!literal) return "NR-15 (MTE), sendo facultada a utilização da NHO 1 da FUNDACENTRO";
                if (literal) return "NR-15 / NHO 1 FUNDACENTRO";
            }
            else if (method == 2) return "NHO 1 da FUNDACENTRO";
            // OUTRA
            
            if (tbruimetodooutro.Text == string.Empty) return "não padronizada";
            return tbruimetodooutro.Text; // 99
        }
        //-----------------------------------------------------
        // POPULA VARIÁVEIS DOS CAMPOS DO FORMULÁRIO DO RUIDO
        //------------------------------------------------ ----
        void popularuido(bool executa)
        {
            if (executa)
            {
                // - Popula o valor aferido da exposição constante do PPP 
                if (tbruiunico.Checked) // está usando valor único na medida? (checa radio button)
                {
                    limite = ledouble(tbruivalue.Text); // pega o limite da caixa de texto 
                }
                else // Não! Portanto está usando valores de valor máximo e mínimo na medida... 
                {
                    limite = min = ledouble(tbmin.Text); // pega limite inferior da caixa de texto 
                    max = ledouble(tbmax.Text); // pega limite superior da caixa de texto 

                }

                unidade = cbbunit.SelectedIndex; // lê a unidade de medida selecionada // 0 = dB(A), 1 = dB
                unidtxt = cbbunit.SelectedItem.ToString();

                // - Popula o método de medição informado no PPP 
                if (rbmetodo.Checked) // Está informando um método "padrão"? (checa radio button)
                {
                    if (cbbmetodo.SelectedIndex == -1) cbbmetodo.SelectedIndex = 7;
                    if (cbbmetodo.SelectedIndex > 2)
                        tbruimetodooutro.Text = cbbmetodo.SelectedItem.ToString();
                    if (cbbmetodo.SelectedIndex == 0) metodo = 1; // NR-15 anexo I
                    else if (cbbmetodo.SelectedIndex == 1) metodo = 12; // NR-15 / NHO 01 FUNDACENTRO
                    else if (cbbmetodo.SelectedIndex == 2) metodo = 2; // NHO 01 FUNDACENTRO
                    else metodo = 99;// Qualquer outro método incorreto  // 99 = outro (inválido) 
                }
                else // RBMETODO OUTRO CHECKED
                {
                    if (tbruimetodooutro.Text.ToUpper().Equals("NR-15")) metodo = 1;
                    else metodo = 99; // Não! Portanto está usando método incorreto 
                }

            }
        }
        // ------------------------------------------------------------------------------------------------------------------------------------------
        // Função Limpatudo - apaga variaveis dinamicas da análise
        // -------------------------------------------------------------------------------------------------------------------------------------------
        void limpatudo()
        {
   // limpa variaveis ruido 
   ruido.limite = ruido.metodo = ruido.tudo = ruido.unidade = ruido.especial = ruido.tudo = true; ;
            ruido.permanencia = cbruidoperma.Checked; 
            popularuido(cbruido.Checked);
            // limpa variaveis eletricidade 
            ele.tudo = true; ele.epoca = true; ele.permanencia = true; ele.limite = true;
            ele.metodo = true;

        }
        // ------------------------------------------------------------------------------------------------------------------------------------------
        // Função Analisa - Analisa o período inicio-fim em relação à janela repassada à função (j.inicio - j.fim) e gera string com laudo discursivo 
        // -------------------------------------------------------------------------------------------------------------------------------------------
        string analisa(janela j)
        {
            limpatudo();
            laudo l; // laudo temporário, com variaveis prazo, resultado, amparo, motivo
            string laudo = ""; // string que vai armazenar o resultado 

            // Algoritmo de análise 
            // 0. Verifica se vai utilizar a janela da lei 
            // 1. Verifica as formalidades (caixas de erros possíveis, no final)                    
            // 2. verifica se a metodologia é válida para o período. 
            // 3. verifica unidade de medida informada 
            // 4. verifica se os valores estão acima do limite 

            // 0. --- Verifica se a data inicial do período está dentro da janela a analisar 
            // Aqui temos tres situações: o periodo começar fora da janela (não usa janela),
            //o período começar na janela anterior (fragmento iniciou no início da janela), 
            //ou o periodo começar dentro da janela (fragmento inicia no início do período). 
            if (DateTime.Compare(inicio, j.fim) > 0 || DateTime.Compare(fim, j.inicio) < 0)  // o periodo inicia após fim ou termina antes do inicio da janela 
            {
                //laudo += "Janela da legislação entre " + j.inicio.ToShortDateString() + " a " + j.fim.ToShortDateString() + " não utilizada. ";
                //laudo += ENTER; // enter 
                return laudo; // sai do método 
            }
            else // está dentro do período. 
            if (DateTime.Compare(inicio, j.inicio) <= 0) l.inicio = j.inicio; // o periodo inicia antes do inicio da janela? sim! Então o inicio será o início da janela 
            else l.inicio = inicio; // não! então o início será o início do período 
            // Agora verifica a data final do periodo em relação a data final da janela. 
            // Aqui temos duas situações: o periodo acaba antes do término da janela (a data final do fragmento é o final do periodo),
            // ou o periodo acaba após a data final da janela, nesse caso a data final do fragmento é a data final da janela.  
            if (DateTime.Compare(fim, j.fim) <= 0) l.fim = fim; // o periodo termina antes do termino da janela?
            else l.fim = j.fim; // não! então concluimos o fragmento ao final da janela, e vamos ter novo fragmento. 
            laudo += l.inicio.ToShortDateString() + " a " + l.fim.ToShortDateString() + ENTER; // Adiciona ao laudo: Período: xx/xx/xx a xx/xx/xx
            laudo += "Normativa vigente: " + j.lei + ". ";

            //------------------------------------------------
            // ANÁLISE DO RUÍDO
            //------------------------------------------------

            if (cbruido.Checked) // SOMENTE SE ESTIVER MARCADA A CHECKBOX PARA RUÍDO
            {

                vesenquadra(j); // verifica se tem requisitos para enquadrar e seta as flags 
                //----------CONCLUSAO: INTRODUÇÃO ------------
                laudo += "Agente: RUÍDO ";
                if (ruido.tudo) laudo += "(cód. " + j.codigo[0] + ") ";
                laudo += "- Conforme Art. 280 da IN 77/2015/INSS, haverá enquadramento " +
                $"quando houver exposição permanente acima de {j.limite.ToString()} dB(A), utilizando-se a técnica " +
                $"{getmetodo(j.metodo)}. Conforme PPP, houve exposição ";
                if (!ruido.permanencia) laudo += "eventual ";

                if (!rbrunaoinformado.Checked) // VERBOSE QUANTITATIVO FOI INFORMADO 
                {
                    laudo += $"a {limite.ToString()}";
                    if (rbruiminmax.Checked) laudo += "-" + max.ToString();
                    laudo += $" {unidtxt}";
                }
                else // QUANTITATIVO NÃO FOI INFORMADO 
                    laudo += "a intensidade sonora não especificada ou desconhecida";
                laudo +=$", utilizando metodologia {getmetodo(metodo, true)}. Conclusão: ";
                // VERIFICA "OUTRO" MOTIVO PARA NÃO ENQUADRAMENTO
                if (cbruneoutro.Checked)
                    if (tbruneoutro.Text == "")
                    {
                        MessageBox.Show("Erro: Não foi especificado motivo para indeferimento do ruído.");
                        cbruneoutro.Checked = false;
                    }
                    else
                        ruido.tudo = false;
                // VERIFICA QUANTITATIVO NÃO INFORMADO
                if (rbrunaoinformado.Checked) ruido.tudo = false;
                //----------CONCLUSOU: ENQUADROU TUDO ------------
                if (ruido.tudo)
                {
                    passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 

                    laudo += " Período ENQUADRADO, considerando que houve exposição " +
                        "permanente, aferida utilizando metodologia correta, acima do limite definido pela legislação.";
                    if (cblayout.Checked) laudo += " Há informação expressa de que não houve alteração no ambiente de trabalho, nos termos "
                            + "do Art. 261 da IN 77/2015.";

                }
                else
                {
                    //----------INICIO DA CONCLUSAO: PERÍODO NÃO ENQUADRADO ------------
                    laudo += " Período NÃO ENQUADRADO. Elementos para não enquadramento: ";
                    //----------CONCLUSAO: MÉTODO INCORRETO ------------
                    if (rbrunaoinformado.Checked) laudo += "Não foi informado ou aferido o quantitativo da exposição sonora. ";
                    if (cbruneoutro.Checked) // indeferir por "outro" motivo
                    {
                        laudo +=tbruneoutro.Text; // adiciona o campo "outro"
                        if (tbruneoutro.Text[tbruneoutro.Text.Length - 1] != '.') // sem ponto no final 
                            laudo += "."; // coloca ponto no final 
                        laudo += " "; // coloca espaço no final 
                    }

                    if (!ruido.permanencia) laudo += NAOPERMA;
                    if (!ruido.metodo)
                    {
                        laudo += "A metodologia utilizada é inválida, sendo " + getmetodo(metodo, true) + ", quando deveria ser utilizada ";
                        if ((metodo == 12) && (j.metodo == 1 || j.metodo == 2) && !cblayout.Checked) laudo += "exclusivamente a ";

                        laudo += getmetodo(j.metodo) + PT;
                        if ((metodo == 12 || metodo == 2) && DateTime.Compare(dbinicio.Value,PUBNHO) < 0 && !cblayout.Checked) // Usou NHO antes de 2001?
                            laudo += " A metodologia NHO 01 só foi publicada pela FUNDACENTRO em 2001, sendo " +
                              "o seu uso anterior a 2001 considerado como inconsistência técnica do PPP, na ausência de observação expressa de que não houve " +
                              "alteração de rotinas de trabalho, maquinário e layout. ";
                        else if ((metodo == 12 || metodo == 2) && j.num <5 && !cblayout.Checked) // Usou NHO antes de 2001?
                            laudo += " A indicação de uso da metodologia NHO 01 em períodos anteriores a 19/11/03 somente pode ser permitida mediante " +
                                "observação expressa de que não houve alteração de rotinas de trabalho, maquinário e layout na empresa, o que não foi constatado. ";
                    }
                    //----------CONCLUSAO: NÃO ULTRAPASSA O LIMITE ------------
                    if (!ruido.limite && !rbrunaoinformado.Checked)
                    {
                        laudo += $" A exposição informada está abaixo do limite necessário para enquadramento, sendo";
                        if (limite == j.limite && !rbruiminmax.Checked) laudo += " exatamente";
                        laudo += $" {limite.ToString()}";
                        if (rbruiminmax.Checked) laudo += "-" + max.ToString();
                        laudo += $" {cbbunit.SelectedItem.ToString()}, quando seria necessário estar ";
                        if (rbruiminmax.Checked) laudo += "permanentemente ";
                        laudo += $"acima de {j.limite.ToString()} dB(A). ";
                    }
                    if (cb_unidade_errada.Checked && !rbrunaoinformado.Checked) // nao aceita unidade
                    {
                        laudo += $"A unidade de medida sonora informada no PPP ({cbbunit.SelectedItem.ToString()}) não é válida, somente sendo " +
                            $"aceita a unidade dB(A) ou NEN. ";
     }
     
		 if (!ruido.especial)
		 {
			laudo += "Valores não expressos em NEN: Mesmo com indicação de uso da metodologia da NHO-01, " +
							 "sem que haja a menção por escrito do uso da NEN nos campos 15.4 ou 15.5 do PPP, " +
							 "não poderá ser aceita, uma vez que dentre as metodologias da NHO-01 encontram-se " +
							 "outras formas de aferição, tais como Leq e TWA. ";
		 }












		}
   }
            //---------
            if (cbeletricidade.Checked) // deve analisar eletrecidade. 
            {

                laudo += "Agente: ELETRICIDADE " + "(cód 1.1.8) - Período ";

                testaele(j);

                if (ele.tudo && ele.permanencia)
                {
                    passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                    laudo += "ENQUADRADO: Conforme Art. 280 da IN 77/2015/INSS, tendo em vista a exposição ocupacional a eletricidade, habitual e permanente acima de 250 volts. ";
                }
                else if (!ele.tudo) // não enquadrou 
                {
                    laudo += "NÃO ENQUADRADO: ";
                    if (!ele.epoca)
                    {
                        laudo += "Após 5 de março de 1997, com a publicação do Decreto nº 2.172, de 1997. este " +
                        "agente foi excluído definitivamente para fins de tempo de serviço como especial. ";
                    }
                    if (!ele.limite)
                    {
                        if (rbeleacima250.Checked) laudo += "A exposição informada como 'acima de 250 volts' não permite avaliar corretamente " +
                                "a efetiva exposição, apesar de convenientemente repetir as palavras da legislação. Devem ser informados valores " +
                                "quantitativos de tensão para permitir o enquadramento. ";
                        else laudo += "Os níveis de exposição não são superiores a 250 volts. ";
                    }
                    if (!cbeleperma.Checked) laudo += NAOPERMA;
                    if (cbeleoutro.Checked && tbeleoutro.Text != string.Empty)
                        laudo += $"{tbeleoutro.Text}. ";
                    if (rbeleunidade.Checked && ele.epoca)
                        laudo += "A unidade de tensão informada não é o 'volt', impossibilitando o enquadramento, que demanda exposição a tensões " +
                            "acima de 250 volts. ";
                    
                }
                if (!ele.metodo)
                {
                    laudo += "Não foram informados os níveis (voltagem) de exposição. ";
                }

            }
            //----------------------------------------------------
            // QUÍMICOS
            //----------------------------------------------------
            if (cbusaq.Checked) // deve analisar quimicos. 
            {
                string naoquant = "NÃO ENQUADRADO: Consta do anexo IV do anexo IV do decreto 3.048/99, porém consta do anexo XI da NR-15 (" +
                                    "agentes cuja insalubridade depende de limite de tolerância) e o PPP não informou exposição " +
                                    "acima dos limites definidos nos anexos XI/XII da NR-15. ";
                string naoquant97 = "NÃO ENQUADRADO: Consta do anexo IV do anexo IV do decreto 2.172/97, porém consta do anexo XI da NR-15 (" +
                                    "agentes cuja insalubridade depende de limite de tolerância) e o PPP não informou exposição " +
                                    "acima dos limites definidos nos anexos XI/XII da NR-15. ";
                string simquali = "ENQUADRADO qualitativamente: Consta dos agentes quimicos analisados qualitativamente por fazer parte " +
                                    "do anexo IV do decreto 3.048/99, e anexo XIII da NR-15, não possuindo limite seguro de tolerância. ";
                string simquali97 = "ENQUADRADO qualitativamente: Consta dos agentes quimicos analisados qualitativamente por fazer parte " +
                                    "do anexo IV do decreto 2.172/97, e anexo XIII da NR-15, não possuindo limite seguro de tolerância. ";
                string simquant = "ENQUADRADO quantitativamente: Consta dos do anexo IV do anexo IV do decreto 3.048/99, estando acima dos" +
                                "limites de tolerância definidos nos anexos XI/XII da NR-15. ";
                string simquant97 = "ENQUADRADO quantitativamente: Consta dos do anexo IV do anexo IV do decreto 2.172/97, estando acima dos" +
                                "limites de tolerância definidos nos anexos XI/XII da NR-15. ";
                string naoIV = "NÃO ENQUADRADO: O elemento não consta do anexo IV do decreto 3.048/99. ";
                
                string naoIV97 = "NÃO ENQUADRADO: O elemento não consta do anexo IV do decreto 2.172/97. ";

                string generico = "NÃO ENQUADRADO: A designação genérica não permite identificar corretamente qual é o elemento químico. ";
                laudo += "Categoria de agente: QUÍMICO - ";
               
                if (lbqfinal.Items.Count > 0) // use metodo completo 
                {
                    for (int i = 0; i < lbqfinal.Items.Count; i++)
                    {
                        

                        laudo += $"Produto: {lbqfinal.Items[i].ToString()} - ";
                        int idx = i+1; // primeiro produto 
                        if (j.num <= 1) // aki
                        // COLORE PRIMEIRA DATA
                        {
                            if (qagflag[idx, 6]) laudo += $"NÃO ENQUADRADO: {NAOPERMA}";  // 6=eventual  
                            else if (qagflag[idx, 7]) laudo += generico;  // 6=eventual 7=generico 
                            else if (qagflag[idx, 0] || qagflag[idx, 1])
                            {
                                laudo += "ENQUADRADO qualitativamente: ";
                                if (qagflag[idx, 0]) laudo += "Faz parte da relação exaustiva do decreto 53.831/64 e 83.080/79. ";
                                if (qagflag[idx, 1]) laudo += "Consta da relação exaustiva da OIT: Model code of safety regulations/1962 -" +
                                        "'Regulamento Tipo Segurança'. ";
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else laudo += "NÃO ENQUADRADO: Não consta dos decretos 53.831/64 e 83.080/79 ou da relação exaustiva da OIT:" +
                                    " Model code of safety regulations/1962 -'Regulamento Tipo Segurança'. ";

                        }
                        else if (j.num > 1 && j.num <5)     // COLORE SEGUNDA
                        {
                            if (qagflag[idx, 6]) laudo += $"NÃO ENQUADRADO: {NAOPERMA}";  // 6=eventual  
                            else if (qagflag[idx, 7]) laudo += generico;  // 6=eventual 7=generico 
                            else if (qagflag[idx, 2]) // quali
                            {
                                if (j.num == 4) laudo += simquali;
                                else laudo += simquali97;
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else if (qagflag[idx, 3] && qagflag[idx, 4]) // 3 = QUANT, 4 = QUANT OK
                            {
                                if (j.num == 4) laudo += simquant;
                                else laudo += simquant97;
                                poelimite();
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else if (qagflag[idx, 3]) // 3 = QUANT, 4 = QUANT OK
                            {
                                if (j.num == 4) laudo += naoquant;
                                else laudo += naoquant97;
                                poelimite();
                            }
                            else
                            {
                                if (j.num == 4) laudo += naoIV;
                                else laudo += naoIV97;
                            }
                        }
                        else if (j.num == 5) // NOVEMBRO 2003
                        {
                            if (qagflag[idx, 6]) laudo += $"NÃO ENQUADRADO: {NAOPERMA}";  // 6=eventual  
                            else if (qagflag[idx, 7]) laudo += generico;  // 6=eventual 7=generico 
                            else if (qagflag[idx, 2]) // quali
                            {
                                laudo += simquali;
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else if (qagflag[idx, 3] && qagflag[idx, 4]) // 3 = QUANT, 4 = QUANT OK
                            {
                                laudo += simquant;
                                poelimite();
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else if (qagflag[idx, 3]) // 3 = QUANT, 
                            {
                                laudo += naoquant;
                                poelimite();
                            }
                            else laudo += naoIV;
                        }
                        // VERIFICA PENULTIMA DATA
                        else if (j.num == 6) // 2004-2014
                        {
                            if (qagflag[idx, 6]) laudo += $"NÃO ENQUADRADO: {NAOPERMA}";  // 6=eventual  
                            else if (qagflag[idx, 7]) laudo += generico;  // 6=eventual 7=generico 
                            else if (qagflag[idx, 2]) // quali
                            {
                                laudo += simquali;
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else if (qagflag[idx, 3] && qagflag[idx, 4]) // 3 = QUANT, 4 = QUANT OK
                            {
                                laudo += simquant;
                                poelimite();
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else if (qagflag[idx, 3]) // 3 = QUANT,  
                                laudo += naoquant;
                            else
                            {
                                laudo += naoIV;
                                poelimite();
                            }
                        }
                        // VERIFICA ULTIMA DATA
                        // 0==decreto 64 1==oit  2==97quali   3==97quant  4==97quantOK  5==LINACH  6==eventual  7==generico
                        // 8==aceitar metodologia   9==NR15   10==NHO

                        else if (j.num == 7) // 2014+
                        { // AKI
                            if (qagflag[idx, 6]) laudo += $"NÃO ENQUADRADO: {NAOPERMA}";  // 6=eventual  
                            else if (qagflag[idx, 7]) laudo += generico;  // 6=eventual 7=generico 

                            else if (qagflag[idx, 5]) // LINACH
                            {
                                laudo += "ENQUADRADO qualitativamente, tendo em vista ser produto comprovadamente cancerígeno " +
                                   "do grupo 1 da LINACH (Lista Nacional de Agentes Cancerígenos para Humanos), pussui registro no" +
                                   " CAS (Chemical Abstracts Service) e consta do anexo IV do Decreto nº 3.048, de 1999," +
                                   "conforme Portaria Interministerial MTE/MS/MPS nº 9, de 2014. ";
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else if (qagflag[idx, 2]) // qualitativo. 
                            {
                                laudo += simquali;
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else if (qagflag[idx, 3] && qagflag[idx, 4]) // 3 = QUANT, 4 = QUANT OK
                            {
                                laudo += simquant;
                                poelimite();
                                passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                            }
                            else if (qagflag[idx, 3]) // 3 = QUANT,  
                            {
                                laudo += naoquant;
                                poelimite();
                            }

                            else laudo += naoIV;

                        }
                        // imprime o limite 
                        void poelimite()
                        {
                            if (lim.ContainsKey(lbqfinal.Items[i].ToString())) // transcreve o valor de referencia 
                                laudo += $"Ref: {lim[lbqfinal.Items[i].ToString()]}. ";
                        }
                    }

                }
                else // usa metodo simplificado 
                {
                    laudo +="Período ";

                    if (rbdaquimi.Checked)
                    {
                        passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                        laudo += "ENQUADRADO: Tendo em vista a exposição ocupacional habitual e permanente a ";
                        foreach (Control c in gbdaquimi.Controls)
                        {
                            if (c is CheckBox)
                            {
                                CheckBox chk = (CheckBox)c;
                                if (chk.Checked)
                                {
                                    if (string.Compare(chk.Text, "Outros:") != 0)
                                    { laudo += chk.Text; laudo += ", "; }
                                }
                            }
                        }
                        if (cbqo.Checked) { laudo += tbqo.Text; laudo += ", "; }
                        laudo += " nos termos da legislação vigente à época, considerando-se a nocividade do agente para a " +
                            "saúde do trabalhador.";


                    }
                    else if (rbnegaquimi.Checked) // não enquadrou 
                    {
                        laudo += "NÃO ENQUADRADO: ";
                        if (cbqoeg.Checked)
                            laudo += "A designação genérica do agente 'óleos e graxas' não permite a caracterização de qual produto se trata, " +
                                 "sendo que existem óleos e graxas nocivos à saúde, e outros altamente purificados que não têm potencial carcinogênico " +
                                 "e podem ser usados inclusive em medicamentos ou cosméticos. ";
                        if (cbqlubrificantes.Checked)
                            laudo += "A designação genérica do agente como 'lubrificantes' não permite a caracterização de qual produto se trata, " +
                                 "sendo que existem óleos lubrificantes nocivos à saúde, e outros inócuos. " +
                                 " Vale ressaltar que a grande maioria dos óleos lubrificantes para motor atualmente utilizados são " +
                                 "do tipo sintético, sem potencial carcinogênico. ";
                        if (cbqom.Checked)
                            laudo += "A designação genérica do agente 'óleo mineral' não permite a caracterização de qual produto se trata, " +
                                 "sendo que existem óleos minerais nocivos à saúde, e outros que não têm potencial carcinogênico " +
                                 "e podem ser usados inclusive em medicamentos ou cosméticos. Vale ressaltar que a grande maioria dos óleos minerais são " +
                                 "sintéticos, portanto sem potencial de nocividade. ";
                        if (cbqs.Checked)
                            laudo += "A designação genérica do agente 'solventes' não permite a caracterização de qual produto se trata, " +
                                 "sendo que existem solventes nocivos à saúde, e outros que não têm potencial carcinogênico " +
                                 "e podem ser usados inclusive em medicamentos ou cosméticos. Deve ser informada a exata composição química do solvente" +
                                 " para possibilitar o enquadramento do período como especial. ";
                        if (cbqpef.Checked)
                            laudo += "A designação genérica do agente 'poeiras e fumos' não permite a caracterização de qual produto se trata, " +
                                 "sendo que existem poeiras e fumos nocivas à saúde, e outras que não têm potencial nocivo ou carcinogênico. " +
                                 "Deve ser informada a exata composição do elemento supostamente nocivo para " +
                                 "possibilitar o enquadramento do período como especial. ";
                        if (cbqh.Checked)
                            laudo += "A designação genérica do agente 'hidrocarbonetos' não permite a caracterização de qual produto se trata, " +
                                 "sendo que existem muitos hidrocarbonetos, alguns nocivos à saúde, e outros que não têm potencial carcinogênico. " +
                                 "Deve ser informada a exata composição dos hidrocarbonetos, se são alifáticos, saturados ou aromáticos, " +
                                 "o tamanho da cadeia ou sua nomenclatura, para possibilitar o enquadramento do período como especial. ";
                        if (cbqhca.Checked)
                            laudo += "A designação genérica 'hidrocarbonetos aromáticos' não permite a caracterização de qual produto se trata, " +
                                 "sendo que existem os que ensejam análise qualitativa (ex: benzeno, antraceno), e outros que necessitam exposição  " +
                                 "acima de limite de tolerância (ex: toluneno - TDI) para efetivo enquadramento. Não foi apresentado laudo técnico das " +
                                 "condições ambientais para se avaliar a composição do produto químico. ";
                        if (cbqotg.Checked)
                            laudo += $"A designação genérica do agente como '{tbqotg.Text}' não permite a caracterização de qual produto se trata, " +
                                 "sendo que essa descrição inclui grande diversidade de agentes, alguns nocivos à saúde, e outros que não têm potencial lesivo. " +
                                 "Deve ser informada a exata composição química dos elemento para possibilitar o enquadramento do período como especial. ";
                        if (cbqppn.Checked)
                        {
                            string S = string.Empty;
                            if (tbqpnn.Text[tbqpnn.Text.Length - 1] == 's') S += 's';
                            laudo += $"O PPP informa exposição a '{tbqpnn.Text}', entretanto não é possível caracterizar a exposição a este{S} agente{S} como período especial, " +
                                 $"haja vista que a nocividade deste{S} elmento{S} é baixa, não gerando por si só risco à saúde do trabalhador, inclusive não estando elencado{S} " +
                                 $"na relação de agentes químicos do anexo IV do Decreto nº 3048/99. ";
                        }
                        if (cbqqal.Checked)
                        {
                            laudo += $"O PPP informa exposição a '{tbqqal.Text}', entretanto não é possível caracterizar esta exposição como período especial, " +
                                 $"haja vista que os níveis de exposição informado no PPP estão inferiores aos limites quantitativos definidos no quadro 8" +
                                 $" do anexo IV do Decreto nº 3048/99 para este produto. ";

                        }
                        if (cbquimieve.Checked) laudo += NAOPERMA;
                    }
                }
            }
            //----------------------------------------------------
            // BIOLÓGICO
            //----------------------------------------------------
            if (cbusabio.Checked) // deve analisar biológicos. 
            {

                // PEGA A PROFISSÃO DO VIVENTE
                //----------------------------
                string prof = string.Empty; // armazena profissão 
                bool semprof = true;
                bool saude = true;
                foreach (Control c in gbbprof.Controls) // corre todas as profissões e encontra a selecionada
                {
                     if (c is RadioButton)
                        {
                            RadioButton chk = (RadioButton)c;
                            if (chk.Checked)
                            {
                            if (string.Compare(chk.Text, "Outra (especificar)") != 0)
                                prof = chk.Text;
                            else 
                                prof = tbboutra.Text; 
                            semprof = false;
                            }
                        }
                }
                prof = prof.ToLower();
                if (!semprof)
                {
                    if (rbboutra.Checked || rbbmotorista.Checked || rbbservgerais.Checked
                    || rbbassistenteadm.Checked || rbbservente.Checked || rbbiosocial.Checked || 
                     rbbiolabo.Checked ||  rbbveterinario.Checked) saude = false;
                }
                laudo += "Categoria de agente: BIOLÓGICOS ";
                if (!rbbnega.Checked)
                {
                    passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                    laudo += "(cód ";
                    if (rbbhumano.Checked) laudo += $"{j.codigo[3]}) ";
                    else laudo += $"{j.codigo2[3]}) ";

                }
                    laudo +="- Período ";
               // PERÍODO NÃO ENQUARADO PARA BIOLÓGICOS
               //--------------------------------------
                if (rbbnega.Checked) // NÃO ENQUADRA 
                {
                    laudo += "NÃO ENQUADRADO: ";
                    // VÊ SE TEM ALGUMA ALTERNATIVA ESPECIFICA DE INDEFERIMENTO MARCADA
                    //------------------------------------------------------------------
                    bool marcounada = true;
                    foreach (Control c in gbbmotivos.Controls) // corre todos os checkbox de motivos
                    {
                        if (c is CheckBox)
                        {
                            CheckBox chk = (CheckBox)c;
                            if (chk.Checked)
                                marcounada = false;
                        }
                    }
                    
                    if (rbbacs.Checked) prof = "agente comunitário de saúde";
                    if (!semprof) laudo += $"Conforme informado no PPP, o requerente exerceu a profissão de {prof} sob risco biológico, entretanto ";
                    if (marcounada)
                    {
                        laudo += "as atividades executadas pelo requerente, conforme descrito no memorial do PPP não envolveram contato " +
                                "direto e permanente com material infectado ou contaminado por patógenos (fungos, bacilos, parasitas, protozoários, vírus, entre outros), nos termos da " +
                                "portaria nº 3.214, de 1978, do MTE.";
                    }
                    else // marcou motivo especifico de indeferimento 
                    {
                        if (cbbeve.Checked) laudo += "não houve exposição permanente aos agentes nocivos (vírus, bactérias, fungos), sendo a " +
                                "exposição realizada de forma eventual ou intermitente, condição necessária para a caracterização da atividade especial; ";
                        if (cbbadm.Checked) laudo += "o serviço executado pelo requerente foi descrito como atividade fundamentalmente administrativa, de forma que " +
                            "não manteve contato com " +
                            "materiais contaminados por patógenos (fungos, bacilos, parasitas, protozoários, vírus, entre outros), nos termos da " +
                            "portaria nº 3.214, de 1978, do MTE; ";
                        if (cbbsemcontato.Checked) laudo += "as atividades executadas pelo requerente, conforme descrito no memorial do PPP não envolveram contato " +
                                "direto com os pacientes / doentes, de forma que não houve efetiva exposição aos agentes nocivos de caráter biológico; ";
                        if (cbbiosembio.Checked) laudo += "as atividades executadas pelo requerente, conforme descrito no memorial do PPP não envolveram contato " +
                                "direto com material infectado por vírus/bactérias/fungos; em que pesem suas características, descaracterizando assim o risco biológico enquadrável; ";
                        laudo += " desqualificando o periodo como 'especial'. ";                        
                    }
                }
                else if (semprof)
                {
                    laudo += "ENQUADRADO: Tendo em vista a exposição ocupacional habitual e permanente a agentes biológicos, pois " +
    "durante o exercício de sua ativdade laboral, conforme informado no PPP, manteve contato com " +
    "materiais contaminados por patógenos (fungos, bacilos, parasitas, protozoários, vírus, entre outros), nos termos da " +
    "portaria nº 3.214, de 1978, do MTE. ";
                }
                else if (saude) // ENQUADRA PROFISSÃO DA SAÚDE 
                {
                    laudo += "ENQUADRADO: Tendo em vista a exposição ocupacional habitual e permanente a agentes biológicos, considerando " +
                        $"ter exercido a profissão de {prof}, com contato com pacientes portadores de doenças infectocontagiosas e/ou manuseio " +
                        "de materiais contaminados por patógenos (fungos, bacilos, parasitas, protozoários, vírus, entre outros), nos termos da " +
                        "portaria nº 3.214, de 1978, do MTE. ";
                }
                else if (!saude) // ENQUADRA PROFISSÃO DISTINTA DA SAÚDE
                {
                    laudo += "ENQUADRADO: Tendo em vista a exposição ocupacional habitual e permanente a agentes biológicos, uma vez " +
                        $"que durante o exercício da profissão de {prof}, apesar de não ser profissão específica da área de saúde, manteve contato com " +
                        "materiais contaminados por patógenos (fungos, bacilos, parasitas, protozoários, vírus, entre outros), nos termos da " +
                        "portaria nº 3.214, de 1978, do MTE. ";
                }
                

            }
            // --------------------------------
            // CALOR
            // ---------------------------------
            if (cbusacalor.Checked)
            {
                bool canega = false;
                bool cada = true;
                bool caneganatural = false;
                bool canegaeventual = false;
                bool cane97quant = false; //  nega quantitativo antes de 97 
                bool cada97qual = false; // concede profissão antes de 97 
                bool cada97quant = false; // concede quantitativo antes  de 97 
                bool caneacima28 = false; // nega quantitativo informado "acima de 28" 
                bool canesemtemp = false; // nega sem temperatura informada ate 97 
                bool cane98semtemp = false; // nega sem temperatura informada apos 97
                bool cane98metodo = false; // nega metodologia incorreta apos 97 
                bool cada98quant = false; // concede quantitativo apos 97 
                bool cane98unidade = false;
                bool cane98quant = false;
                bool canenr15depois = false;
                bool canenhoantes = false; // nega usou NHO 06 antes de 203
                double temp = 0; // quantitativo informado 
                int cametodo = 99;
                if (rbcavalorunico.Checked) temp = ledouble(tbcaunico.Text);
                else if (rbcaminmax.Checked) temp = ledouble(tbcamin.Text);
                else if (rbcaacimade28.Checked && !cbcanegaacimade28.Checked) temp = 28.1;
                

                laudo += "Categoria de agente: CALOR - Período ";
                if (!cbcaeventual.Checked) canegaeventual = true;
                // análise anterior a 1997---------------------------
                
                if (j.antigo)
                {
                    // FONTE NÃO ARTIFICIAL 
                    if (!cbcaartificial.Checked) caneganatural = true;
                    // ENQUADRA QUALITATIVO
                    if (cbcaqualitativo.Checked) // Profissão de enquadramento qualitativo 
                        cada97qual = true;
                    else // antes de 1997, não tem profissao qualitativa, verifica temperatura acima de 28 graus
                    { // não era obrigatório metodologia e nem a unidade correta
                        
                        if (temp > 28) cada97quant = true;
                        else cane97quant = true;
                        if (rbcanaoinformado.Checked) canesemtemp = true;
                        if (rbcaacimade28.Checked && cbcanegaacimade28.Checked) caneacima28 = true;
                    }
                }   
                // análise posterior a 1997---------------------------
                else
                {
              
                    //------------------------------------------------
                    if (rbcaacimade28.Checked && cbcanegaacimade28.Checked) caneacima28 = true;
                    if (rbcanaoinformado.Checked) cane98semtemp = true;
                    // - Popula o método de medição informado no PPP 
                    if (rbcametodo.Checked) // Está informando um método "padrão"? (checa radio button)
                    {
                        if (cbbcametodo.SelectedIndex == -1) cbbcametodo.SelectedIndex = 3;
                        //tbcametodo.Text = cbbcametodo.SelectedItem.ToString();
                        if (cbbcametodo.SelectedIndex == 0) cametodo = 1; // NR-15  
                        else if (cbbcametodo.SelectedIndex == 1) cametodo = 12; // NR-15 / NHO 01 FUNDACENTRO
                        else if (cbbcametodo.SelectedIndex == 2) cametodo = 2; // NHO 06 FUNDACENTRO
                        
                    }
                    // VERIFICA UNIDADE
                    if (cbbcaun.SelectedIndex != 1) cane98unidade = true; // Não informou em IBUTG
                    // VERIFICA METODOLOGIA
                    if (j.metodo != cametodo || (cametodo==99 && !cbcaaceitametodo.Checked))
                    {
                        if (!(j.metodo == 12 && (cametodo == 1 || cametodo == 2)))
                            cane98metodo = true;

                        // ACEITA NHO 06 MESMO ANTES DE 2003
                        if (cbcaextemporaneo.Checked && (cametodo == 2)) cane98metodo = false;
                        if (cbcaextemporaneo.Checked && (cametodo == 12 && j.metodo != 2)) cane98metodo = false;
                        if (j.metodo == 2 && (cametodo == 1 || cametodo == 12)) { canenr15depois = true; cane98metodo = true; }
                        if ((cametodo == 2 || cametodo == 12) && j.metodo == 1 && !cbcaextemporaneo.Checked)
                        { canenhoantes = true; cane98metodo = true; }
                        if (cbcaaceitametodo.Checked && cametodo == 99) cane98metodo = false;
                    }
                    // FAZ ANÁLISE QUANTITATIVA APÓS 97
                    limite = double.Parse(tbcalim.Text);
                    if (temp > limite) cada98quant = true;
                    else cane98quant = true;

                }

                // VERIFICA AS FLAGS
                if (caneganatural || canegaeventual || cane97quant || canesemtemp ||
                    caneacima28 || cane98unidade || cane98metodo || canenhoantes
                    || canenr15depois || cane98semtemp || cane98quant)
                    { canega = true; cada = false; }
                
                if (canega)
                {
                    laudo += "NÃO ENQUADRADO. Elementos para não enquadramento: ";
                    
                    if (caneganatural && j.antigo) laudo += "O anexo do Decreto nº 53.831/64 somente estabelece como especial a operação em locais " +
                    "com temperatura excessivamente alta proveniente de fontes artificiais, não se considerando o calor oriundo do sol, " +
                    "fontes geotérmicas ou aglomerações humanas como enquadrável. ";
                    if (cane97quant && !canesemtemp && !caneacima28) laudo += "O requerente exerceu de atividade profissional distinta das especificadas no " +
                            "Anexo II do Decreto nº 83.080/79 (indústria metalúrgica/mecânica, fornos e caldeiras, manipulação de metais " +
                            "quentes ou fabricação de vidros e cristais) sob temperatura inferior a 28º C, limite constante do anexo do Decreto nº 53.831/64. ";
                    if (caneacima28) laudo += "A informação constante no PPP de trabalho 'acima de 28 graus', apesar de convenientemente repetir o texto da legislação, " +
                            "não permite saber quais as reais temperaturas de operação, impedindo o efetivo enquadramento do período como especial. ";
                    if (canesemtemp) laudo += "O requerente exerceu de atividade profissional distinta das especificadas no " +
                            "Anexo II do Decreto nº 83.080/79 (indústria metalúrgica/mecânica, fornos e caldeiras, manipulação de metais " +
                            "quentes ou fabricação de vidros e cristais) e não há informação de temperatura de trabalho no PPP, a qual é necessária para análise quantitativa " +
                            "prevista no anexo do Decreto nº 53.831/64, com limite mínimo de de 28 graus oriundos de fontes artificiais de calor. ";
                    if (canegaeventual) laudo += NAOPERMA;
                    if (cane98unidade) laudo += "Não foi utilizada a unidade de medida IBUTG, obrigatória a partir de 06/03/1997. ";
                    if (cane98semtemp) laudo += "Não foi informada a temperatura de exposição, impossibilitando a análise quantitativa. ";
                    if (cane98metodo)
                    {
                        string mcerto = string.Empty;
                        if (j.metodo == 1)
                        {
                            if (cametodo == 12) mcerto = "exclusivamente a normativa ";
                            mcerto += "NR-15 do MTE";
                        }
                        else if (j.metodo == 12) mcerto = "NR-15 (MTE), sendo facultada a utilização da NHO 6 da FUNDACENTRO";
                        else if (j.metodo == 2)
                        {
                            if (cametodo == 12) mcerto = "exclusivamente a normativa ";
                            mcerto += "NHO 6 da FUNDACENTRO";
                        }
                        
                        laudo += $"Não foi utilizada a metodologia obrigatória para o período, {mcerto}. ";
                        if (canenhoantes) laudo += "É vedado o uso da metodologia NHO 06 em período anterior a 19/11/2003 sem informação expressa de que não houve" 
                          +" alteração no layout, maquinário, ou rotinas de trabalho, nos termos do Art. 261 da IN 77/2015. ";
                        if (canenr15depois) laudo += "O uso da metodologia NR-15 é vedado após 01/01/2004. ";
                    }
                    if (cane98quant)
                    {
                        laudo += $"O requerente exerceu atividade profissional submetido a carga térmica ";
                        if (!rbcanaoinformado.Checked)
                        {
                            laudo += $" de { temp.ToString()} ";
                            if (cbbcaun.SelectedIndex != 2)
                                laudo += $"{cbbcaun.SelectedItem.ToString()}, ";
                            else laudo += "graus, ";
                        }
                        else
                            laudo += "desconhecida, portanto ";
                                laudo += $"não superando o limite de tolerância para o período, de {tbcalim.Text} IBUTG, definido nos quadros nº 1 " +
                                $"e 2 do anexo 3 da NR-15, considerando-se metabolismo médio ponderado calculado em {tbcaibutg.Text} cal/h, conforme quadro nº 3 do anexo " +
                                $"3 da NR-15.";
                    }

                }
                if (cada)
                {
                    passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                    laudo += "ENQUADRADO ";
                    if (cada97qual) laudo += "(cod. 1.1.1): O requerente exerceu de forma não eventual e permanente atividade profissional discriminada no Anexo II do Decreto " +
                            "nº 83.080, de 1979: Atividades na indústria metalúrgica/mecânica, fornos e caldeiras, manipulação de metais quentes ou fabricação " +
                            "de vidros e cristais.";
                    if (cada97quant) laudo += "(cod. 1.1.1): O requerente exerceu de forma não eventual e permanente atividade profissional distinta das exemplificadas no " +
                            "Anexo II do Decreto nº 83.080/79, sob temperatura superior a 28º C, conforme anexo do Decreto nº 53.831/64. ";
                    if (cada98quant) laudo += $"(cod. 2.0.4): O requerente exerceu de forma não eventual e permanente atividade profissional submetido a carga térmica de {temp.ToString()} IBUTG, " +
                            $"com consequente risco potencial de dano à sua saúde, acima do limite de tolerância de {tbcalim.Text} IBUTG, definido nos quadros nº 1 " +
                            $"e 2 do anexo 3 da NR-15, considerando-se metabolismo médio ponderado calculado em {tbcaibutg.Text} cal/h, conforme quadro nº 3 do anexo " +
                            $"3 da NR-15.";

                

                }
                if (cbcaextemporaneo.Checked) laudo += " Há informação expressa de que não houve alteração no ambiente de trabalho, nos termos "
        + "do Art. 261 da IN 77/2015, permitindo laudo extemporâneo e uso da metodologia NHO 01 anteriormente a 19/11/2003. ";
            }
            // ---------------------------
            // RADIAÇÕES NÃO IONIZANTES
            //----------------------------
            if (cbrni.Checked) // deve analisar radiações não ionizantes
            {
                bool semprof=false;
                bool negarni = false;
                
                // ---------------------------
                // PEGA A PROFISSÃO DO VIVENTE
                //----------------------------
                string prof = string.Empty;
                if (rbrnisoldador.Checked) prof = "soldador";
                else if (rbrniaeroviario.Checked) prof = "aeroviário";
                else // if (rbrniprofoutra.Checked)
                {
                    if (tbrniprof.Text==string.Empty)  prof = "trabalho descrita no campo 14.2 do PPP";  // não disse qual "outra"
                    else prof = tbrniprof.Text; // pega a outra da caixa de texto
                }
                prof = prof.ToLower();
                // PEGA O TIPO DA ONDA
                //----------------------------
                string onda = "";
                foreach(Control c in gbrnitipo.Controls) // corre todos os checkbox de motivos
                {
                    if (c is CheckBox)
                    {
                        CheckBox chk = (CheckBox)c;
                        if (chk.Checked)
                        {
                            if (onda != "") onda += " / ";
                            string s = chk.Text;
                            if (s[1] != 'u') onda += chk.Text.ToLower(); // não é outros
                            else
                            {
                                if (tbrniagenteoutros.Text == string.Empty) onda += "radiações não ionizantes";
                                else onda += tbrniagenteoutros.Text.ToLower(); // outros 
                            }
                        }
                    }
                }
                if (onda == string.Empty) onda = "agentes nocivos";
                // FAZ A ANALISE
                if (rbrninega.Checked) negarni = true;
                if (!j.antigo) negarni = true;
                // VE SE MARCOU CHECKBOX DE INDEFERIR 

                bool rninegabox = false;
                foreach (Control c in gbrnimotivos.Controls) // corre todos os checkbox de motivos
                {
                    if (c is CheckBox)
                    {
                        CheckBox chk = (CheckBox)c;
                        if (chk.Checked)
                            rninegabox = true;
                    }
                }
                //------------------------------
                laudo += "Categoria de agente: RADIAÇÃO NÃO IONIZANTE (1.1.4): ";
                // CABEÇALHO
                if (!semprof && j.antigo) laudo += $"O PPP informa exercício da atividade de {prof} sob ação de {onda}. Período ";
                // PERÍODO NÃO ENQUARADO PARA RADIAÇÃO NÃO IONIZANTES
                //--------------------------------------
                if (negarni)
                {
                    laudo += "NÃO ENQUADRADO. Elementos para não enquadramento: ";
                    if (!j.antigo)
                    {
                        laudo += "Após 5 de março de 1997, com a publicação do Decreto nº 2.172, de 1997. este " +
                        "agente foi excluído definitivamente para fins de tempo de serviço como especial. ";

                    }
                  
                        // VÊ SE TEM ALGUMA ALTERNATIVA ESPECIFICA DE INDEFERIMENTO MARCADA
                        //------------------------------------------------------------------

                        // NÃO MARCOU NADA... INDEFERIMENTO GENÉRICO 
                        else if (!rninegabox)
                        {
                            laudo += "As atividades executadas pelo requerente, conforme descrito no memorial do PPP, no campo 15.2 (profissiografia)" +
                            " não envolveram exposição permanente e não eventual a radiações não ionizantes, desqualificando o periodo como 'especial'. ";
                        }
                        // MARCOU OPÇÃO DE INDEFERIMENTO. UTILIZAR. 
                        else // if (rninegabox)
                        {
                            if (cbrninegasol.Checked) laudo += "A simples atividade laboral sob efeito do sol não pode ser considerada como permanente, já que " +
                                    "durante muitos dias do ano ocorre chuva, tempo nublado ou escuro, quando os raios ultra-violetas solares são bloqueados pelas nuvens, " +
                                    "não havendo assim a permanência";
                            if (cbrninegaeventual.Checked) laudo += NAOPERMA;
                            if (cbrninegaadm.Checked) laudo += "o serviço executado pelo requerente foi descrito como atividade fundamentalmente administrativa (serviço interno e " +
                                    "não operacional), de forma que não foi efetivamente exposto às radiações não-ionizantes,";
                            if (cbrnioutronao.Checked) laudo += $"O agente ao qual o trabalhador foi exposto ({onda}) não é uma radiação " +
                                    $"não ionizante, em que pesem suas características, ";
                            if (cbrninegadistante.Checked) laudo += $"As atividades executadas pelo requerente se situavam demasiadamente distantes das fontes " +
                                    $"de emissões de radiações, de forma que a efetiva exposição foi mínima ou mesmo nula, ";
                            if (cbrninegadesprezivel.Checked) laudo += $"As quantidades de radiação efetivamente recebidas pelo obreiro foram mínimas ou insignificantes " +
                                    $"dadas as características de sua atividade laboral, conforme descritivo na profissiografia do PPP, ";
                            if (cbrninegaoutro.Checked) laudo += tbrninegaoutro.Text + ",";
                            else laudo += " desqualificando o periodo como 'especial'. ";
                        }
                    

                }
                else // ENQUADRA  
                {
                    passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                    laudo += "ENQUADRADO: ";
                    if (rbrnisoldador.Checked) laudo += "O anexo do Decreto 53831/64 cita especificamente a atividade de soldador com arco elétrico e com oxiacetilenio" +
                            " como exposição a radiações não ionizantes, tendo em vista as radiações ultra-violetas emitidas pelos aparelhos de solda, portanto" +
                            "é devido o enquadramento especial até 05/03/97. ";
                    else if(rbrniaeroviario.Checked) laudo += "O anexo do Decreto 53831/64 cita especificamente a atividade de especificamente os aeroviários de manutenção " +
                            "de aeronaves e motores, turbo-hélices e outros, tendo em vistas as radiações não ionizantes emitidas pelas ondas de radar nos ambientes "+
                           "de aeroportos, portanto é devido o enquadramento especial até 05/03/97. ";
                    else 
                            laudo += $" Tendo em vista a exposição ocupacional habitual e permanente a radiação não ionizante ({onda}) durante o exercício  " +
                        $"da atividade de {prof}.";

                }
                
               
            }

            if (cbergo.Checked)
            {
                laudo += $"Categoria de agente: ERGONÔMICO - Período não enquadrado; apesar de constar do PPP, não há previsão legal de caracterização desta espécie de agente como atividade especial. ";
            }
                if (cbfis.Checked)
            {
                laudo += $"Categoria de agente: RISCO DE ACIDENTES - Período não enquadrado; apesar de constar do PPP, não há previsão legal de caracterização desta espécie de agente como atividade especial. ";
            }
            if (cberg2.Checked)
            {
                string agente = tbergo2.Text.ToUpper();
                laudo += $"Categoria de agente: {agente} - Período não enquadrado; apesar de constar do PPP, não há previsão legal de caracterização desta espécie de agente como atividade especial. ";
            }
            // ---------------------------
            //            FRIO
            //----------------------------
            if (cbfrio.Checked)
            {
                //------------------------------
                laudo += "Categoria de agente: FRIO: ";
                // CABEÇALHO

                if (!j.antigo)
                {
                    laudo += "Período NÃO ENQUADRADO. Elementos: Após 5 de março de 1997, com a publicação do Decreto nº 2.172/97 " +
                        "este agente foi excluído definitivamente para fins de tempo de serviço como especial. ";
                }
                else if (cbfrioperma.Checked)
                    laudo += "Período NÃO ENQUADRADO. Elementos: " + NAOPERMA;
                else if (cbfrioadministrativo.Checked)
                {
                    laudo += "Período NÃO ENQUADRADO. Elementos: As atividades exercidas pelo trabalhador conforme profissiografia " +
                        "constante no campo 15.2 do PPP foram essencialmente admistrativas, sem exposição efetiva ao frio. ";
                    if (cbfrifazgelo.Checked) laudo += "Não basta a empresa possuir câmara fria ou mesmo atuar em fabricação de gelo, " +
                            "é necessário que o trabalhador tenha sido efetivamente exposto, de forma permanente e não eventual a " +
                            "temperatura excessivamente baixa, capaz de ser nociva à saúde. ";
                }
                else if (cbfriosemfrio.Checked)
                {
                    laudo += "Período NÃO ENQUADRADO. Elementos: As atividades exercidas pelo trabalhador conforme profissiografia " +
                        "constante no campo 15.2 do PPP não envoveram contato direto com os ambientes frios. ";
                    if (cbfrifazgelo.Checked) laudo += "Não basta a empresa possuir câmara fria ou mesmo atuar em fabricação de gelo, " +
                            "é necessário que o trabalhador tenha sido efetivamente exposto, de forma permanente e não eventual a " +
                            "temperatura excessivamente baixa, capaz de ser nociva à saúde. ";
                }
                else
                {
                    
                    double temp = 13;
                    if (rbfrivalorunico.Checked)
                        temp = ledouble(tbfriunico.Text);
                    else if (rbfriminmax.Checked)
                        temp = Math.Max(ledouble(tbfrimin.Text), ledouble(tbfrimax.Text));
                    else if (rbfriinferiora12.Checked && !cbfrinaoaceita.Checked)
                        temp = 11;
                    var frioquali = new DateTime(1996, 10, 14); // expire date 
                    if (DateTime.Compare(frioquali, j.fim) > 0 && cbfrifazgelo.Checked)
                    {
                        passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                        laudo += "cód 1.1.2 - Período ENQUADRADO qualitativamente, por tipo de atividade profissional. Elementos: O Dec. nº 53.831/64, " +
                            "estabeleceu como especiais as operações " +
                   "em locais com temperatura excessivamente baixa, inferior a 12º C, proveniente de fontes artificiais, capaz " +
                    "de ser nociva à saúde, com jornada normal. O Decreto nº 611, de 1992, permitia também análise qualitativa, " +
                   "nas atividades em câmaras frias e fabricação de gelo, conforme Anexo I do Dec. nº 83.080/79, até 13/10/96. ";
                    }
                    else if (temp < 12 && cbfrifonteartifical.Checked)
                    {
                        passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                        laudo += "cód 1.1.2 - Período ENQUADRADO quantitativamente. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                            "em locais com temperatura excessivamente baixa, inferior a 12º C, proveniente de fontes artificiais, capaz " +
                            "de ser nociva à saúde, com jornada normal. O Decreto nº 611, de 1992, permitia também análise qualitativa, " +
                            "nas atividades em câmaras frias e fabricação de gelo, conforme Anexo I do Dec. nº 83.080/79, até 13/10/96. " +
                            $"No presente requerimento a temperatura informada foi ";
                        if (rbfriinferiora12.Checked) laudo += "inferior a 12º C";
                        else laudo += $"de { temp.ToString()} º C";
                        laudo += ", portanto enquadrável. ";
                    }
                    else if (!cbfrifazgelo.Checked || (DateTime.Compare(frioquali, j.fim) <= 0 && cbfrifazgelo.Checked))
                    {
                        if (!cbfrifonteartifical.Checked)
                            laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                            "em locais com temperatura excessivamente baixa, inferior a 12º C, proveniente de fontes ARTIFICIAIS, capaz " +
                            "de ser nociva à saúde, com jornada normal. O Decreto nº 611, de 1992, permitia também análise qualitativa, " +
                            "nas atividades em câmaras frias e fabricação de gelo, conforme Anexo I do Dec. nº 83.080/79, até 13/10/96. " +
                            "No presente requerimento, conforme profissiografia constante no campo 15.2 do PPP, temos que a origem da " +
                            "fonte de frio é NATURAL, portanto não enquadrável para período especial. ";
                        else if (rbfrisemquant.Checked)
                            laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                            "em locais com temperatura excessivamente baixa, inferior a 12º C, proveniente de fontes artificiais, capaz " +
                            "de ser nociva à saúde, com jornada normal. O Decreto nº 611, de 1992, permitia também análise qualitativa, " +
                            "nas atividades em câmaras frias e fabricação de gelo, conforme Anexo I do Dec. nº 83.080/79, somente até 13/10/96. " +
                            "No presente requerimento, não foi informado o valor das temperaturas ao qual foi exposto, impossibilitando " +
                            "a análise quantitativa. ";
                        else if (rbfriinferiora12.Checked && cbfrinaoaceita.Checked)
                            laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                            "em locais com temperatura excessivamente baixa, inferior a 12º C, proveniente de fontes ARTIFICIAIS, capaz " +
                            "de ser nociva à saúde, com jornada normal. O Decreto nº 611, de 1992, permitia também análise qualitativa, " +
                            "nas atividades em câmaras frias e fabricação de gelo, conforme Anexo I do Dec. nº 83.080/79, somente até 13/10/96. " +
                            "No presente requerimento, conforme profissiografia constante no campo 15.2 do PPP, a temperatura " +
                            "de exposição foi descrita apenas como 'inferior a 12' graus. Para possibilitar o enquadramento seria necessário que" +
                            "a temperatura fosse aferida e informada precisamente em graus celcius, e não de forma vaga, convenientemente " +
                            "repetindo o texto da legislação. ";
                        else if (temp >= 12)
                            laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                            "em locais com temperatura excessivamente baixa, inferior a 12º C, proveniente de fontes artificiais, capaz " +
                            "de ser nociva à saúde, com jornada normal. O Decreto nº 611, de 1992, permitia também análise qualitativa, " +
                            "nas atividades em câmaras frias e fabricação de gelo, conforme Anexo I do Dec. nº 83.080/79, somente até 13/10/96. " +
                            $"No presente requerimento, a temperatura de trabalho informada foi de {temp.ToString()}º C, portanto não inferior " +
                            $"a 12º C, descaracterizando o enquadramento quantitativo. ";
                    }
                }



            }
            // ---------------------------
            //            UMIDADE
            //----------------------------
            if (cbumidade.Checked)
            {
                //------------------------------
                laudo += "Categoria de agente: UMIDADE: ";
                // CABEÇALHO

                if (!j.antigo)
                {
                    laudo += "Período NÃO ENQUADRADO. Elementos: Após 5 de março de 1997, com a publicação do Decreto nº 2.172/97 " +
                        "este agente foi excluído definitivamente para fins de tempo de serviço como especial. ";
                }
                else if (j.num == 1 && cbumisemltcat.Checked)
                {
                    laudo += "Período NÃO ENQUADRADO. De 14/10/1996 a 05/03/1997 passaram a ser exigidos LTCAT ou outras " +
                        "demonstrações ambientais para efetiva comprovação da insalubridade, mediante inspeção realizada no local de trabalho. " +
                        "Não constam no presente processo estes documentos essenciais ao enquadramento do período como especial. ";
                }
                else if (rbumiconcede.Checked)
                {
                    passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                    laudo += "cód 1.1.3 - Período ENQUADRADO qualitativamente, tendo em vista exercício de atividade laboral em local " +
                        "em contato direto e permanente com água e umidade excessiva, capaz de ser nociva à saúde e proveniente de " +
                        "fonte artificial, de forma habitual e permanente, conforme Anexo 10 da NR-15. ";
                }
                else if (cbuminaoperma.Checked)
                    laudo += "Período NÃO ENQUADRADO. Elementos: " + NAOPERMA;
                else // indefere por outros motivos 
                {
                    if (cbumiadministrativo.Checked)
                    {
                        laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                       "em locais com umidade excessiva, em contato direto e permanente com água, capazes de serem nocivas à saúde " +
                       "e proveniente de fontes artificiais. Entretanto, as atividades exercidas pelo trabalhador conforme profissiografia " +
                            "constante no campo 15.2 do PPP foram essencialmente admistrativas, sem exposição efetiva à umidade. ";
                    }
                    else if (cbumisemumi.Checked)
                    {
                        laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                       "em locais com umidade excessiva, em contato direto e permanente com água, capazes de serem nocivas à saúde " +
                       "e proveniente de fontes artificiais. Entretanto, em revisão aos documentos existentes no processo, bem como às " +
                       "atividades exercidas pelo trabalhador descritas na profissiografia constante no campo 15.2 do PPP, verificamos que " +
                       "o labor não envolveu contato direto e permanente com água nem foi realizado em locais alagados ou encharcados.";

                    }
                    else if (rbuminatural.Checked)
                        laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                        "em locais com umidade excessiva, em contato direto e permanente com água, capazes de serem nocivas à saúde " +
                        "e proveniente de fontes ARTIFICIAIS. No presente requerimento, conforme profissiografia constante no campo 15.2 " +
                        "do PPP, temos que a origem da umidade foi NATURAL, portanto não enquadrável para período especial. ";
                    else if (rbumiambienteseco.Checked)
                    {
                        laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                        "em locais com umidade excessiva, em contato direto e permanente com água, capazes de serem nocivas à saúde " +
                        "e proveniente de fontes artificiais. Entretanto, em revisão aos documentos existentes no processo, bem como às " +
                        "atividades exercidas pelo trabalhador descritas na profissiografia constante no campo 15.2 do PPP, verificamos que o " +
                        "ambiente de trabalho era prevalentemente seco, não gerando contato direto e contínuo com a umidade. ";
                    }
                    else if (rbumioutro.Checked)
                    {
                        if (tbumioutro.Text == string.Empty) laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                       "em locais com umidade excessiva, em contato direto e permanente com água, capazes de serem nocivas à saúde " +
                       "e proveniente de fontes artificiais. Entretanto, em revisão aos documentos existentes no processo, bem como às " +
                       "atividades exercidas pelo trabalhador descritas na profissiografia constante no campo 15.2 do PPP, verificamos que " +
                       "o labor não envolveu contato direto e permanente com água nem foi realizado em locais alagados ou encharcados.";
                        else laudo += "Período NÃO ENQUADRADO. Elementos: O Dec. nº 53.831/64, estabeleceu como especiais as operações " +
                       "em locais com umidade excessiva, em contato direto e permanente com água, capazes de serem nocivas à saúde " +
                       "e proveniente de fontes artificiais. Entretanto, em revisão aos documentos existentes no processo verificamos que: " + tbumioutro.Text;
                    }
                }
            }
            // ---------------------------
            //            VIBRAÇÃO
            //----------------------------
            if (cbvibra.Checked)
            {
                //------------------------------
                laudo += "Categoria de agente: VIBRAÇÃO: ";
                // CABEÇALHO

                if (cbvmartelete.Checked && !cbvnaoperma.Checked)
                {
                    passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                    laudo += "Período ENQUADRADO, tendo em vista o exercício de trabalho com perfuratrizes ou marteletes pneumáticos. ";
                }
                else if (cbv120.Checked && j.antigo && !cbvnaoperma.Checked)
                {
                    passou[j.num] = true; // seta flag do periodo aprovado, para verificar depois segmentação dos periodos 
                    laudo += "Período ENQUADRADO, tendo em vista o exercício de trabalho com máquinas acionadas por ar comprimido e " +
                        "velocidade acima de 120 golpes por minuto, conforme Art. 187 da CLT e Portaria Ministerial 262, de 06/08/1962. ";
                }
                else
                {
                    laudo += "Período NÃO ENQUADRADO. Elementos: ";
                    if (cbvnaoperma.Checked) laudo += NAOPERMA;
                    else
                    {
                        laudo += "Não exerceu atividade de trabalho com perfuratrizes pneumáticas ou marteletes pneumáticos";
                        if (j.antigo) laudo += " bem como não laborou operando máquinas acionadas por ar comprimido e velocidades acima de 120 golpes por minuto, " +
                            "nos termos do Art. 187 da CLT e Portaria Ministerial 262, de 06/08/1962";
                        laudo += ". ";
                    }

                }

                        
            }
            laudo += ENTERS;
            return laudo;
        }

        



        // -------------------------------------------------
        // CONFIGURA AÇÕES DE BOTÕES E ELEMENTOS INTERATIVOS  
        // -------------------------------------------------
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbbunit.SelectedIndex != 0) cb_unidade_errada.Checked = true; // caso selecionar unidade esquisita, marca "não aceitar" por padrão
            else cb_unidade_errada.Checked = false;
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {
            
        }

        private void toolTip1_Popup_1(object sender, PopupEventArgs e)
        {

        }

        private void cbunidade_CheckedChanged(object sender, EventArgs e)
        {
           //if (cbunit.Checked) cbbunit.SelectedIndex = 1; // se marcar "não aceitar unidade", muda para db
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (rbruiminmax.Checked) // selecionou minimo / máximo 
            {
                tbruivalue.Enabled = false;
                tbmin.Enabled = tbmax.Enabled = true;
            }
        }

        private void rbunico_CheckedChanged(object sender, EventArgs e)
        {
            if (tbruiunico.Checked) // selecionou minimo / máximo 
            {
                tbruivalue.Enabled = true;
                tbmin.Enabled = tbmax.Enabled = false;
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (rbruimetodooutro.Checked)
            {
                tbruimetodooutro.Enabled = true;
                cbbmetodo.Enabled = false;
            }

        }

        private void rbmetodo_CheckedChanged(object sender, EventArgs e)
        {
            if (rbmetodo.Checked)
            {
                tbruimetodooutro.Enabled = false;
                cbbmetodo.Enabled = true;
            }
        }

        private void toolTip3_Popup(object sender, PopupEventArgs e)
        {

        }

        // ------------------------------
        // Clica botão avaliar 
        // ------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                checkexpire();
                if (dbfim.Value < new DateTime(1964, 01, 01))
                {
                    fala("Erro: Período anterior a 1964.");
                    return;
                }
                if (DateTime.Compare(dbinicio.Value, dbfim.Value) > 0)
                {
                    fala("A data de início não pode ser anterior à data de fim.");
                    return;
                }
                for (int i = 0; i < passou.Count; i++) passou[i] = false; // reseta flag passou 
                if (!cbeletricidade.Checked && !cbruido.Checked && !cbusaq.Checked && !cbusabio.Checked
                    && !cbrni.Checked && !cbusacalor.Checked && !cbergo.Checked && !cberg2.Checked
                    && !cbfis.Checked && !cbfrio.Checked && !cbumidade.Checked && !cbvibra.Checked)
                {
                    MessageBox.Show("Erro: nenhum agente nocivo selecionado.");
                    return;
                }
                //try
                //{
                tblaudo.Text = ""; // limpa quadro
                                   // ------------------------------------------------------------
                                   // Popula VARIÁVEIS dinâmicas com valores obtidos do formulário 
                                   // ------------------------------------------------------------
                inicio = dbinicio.Value; // Pega data de início da caixa de texto 
                fim = dbfim.Value; // Pega data final da caixa de texto 
                                   // ------------------------------------------------------------
                                   // Analisa o período uma por uma as janelas da legislação 
                                   // ------------------------------------------------------------
                if (tbouprofissaoppp.Text != "") tblaudo.Text += $"Profissão (PPP): {tbouprofissaoppp.Text}{ENTER}";
                if (tbouprofctps.Text != "") tblaudo.Text += $"Profissão (CTPS): {tbouprofctps.Text}{ENTER}";
                // ---------------------------------------------
                // GERA O TEXTO DAS ANALISES DOS PERÍODOS 
                //----------------------------------------------
                var js = new janela[8] { j2, j3, j4, j5, j6, j7, j8, j9 }; // adiciona as janelas a um array 
                int red = 1; // redutor para parar antes do periodo de quimicos 2014 - linach 
                if (cbusaq.Checked && DateTime.Compare(dbfim.Value, j9.inicio) >= 0) // marcou quimico e passou de 2014
                {
                    red = 0; // elimina redutor 
                    js[6] = j8b; // substitui o ultimo periodo usual por periodo intermediario para acessar 08/10/14+
                    if (passou.Count < 8) passou.Add(false);
                }
                else
                {
                    if (passou.Count == 8) passou.RemoveAt(7);
                }
                for (int i = 0; i + red < js.Length; i++) // iterate periodos fazendo a análise 
                    tblaudo.Text += analisa(js[i]); //  adiciona texto a caixa de texto

                // ---------------------------------------------
                // VERIFICA SE É NECESSÁRIO FRACIONAR PERÍODOS 
                //----------------------------------------------

                var seg = new List<List<int>>();
                bool last = passou[0];
                seg.Add(new List<int>()); // inicializa a lista seg. 

                for (int i = 0; i < passou.Count; i++)
                { // criando grupos de periodos com valores semelhantes. 
                    if (DateTime.Compare(js[i].inicio, dbfim.Value) <= 0 &&
                        DateTime.Compare(js[i].fim, dbinicio.Value) >= 0)
                    {
                        if (seg[0].Count == 0) last = passou[i];
                        if (passou[i] == last) seg[seg.Count - 1].Add(i); // se tá igual ao ultimo resultado...  adiciona o periodo
                        else // ficou diferente -> vai ter  qeu fragmentar
                        {
                            last = passou[i];
                            var temp = new List<int>(); // cria novo grupo 
                            temp.Add(i); // adiciona o periodo a um novo grupo 
                            seg.Add(temp); // adiciona o grupo à lista de periodos a se fragmentar. 
                        }
                    }
                }

                // ---------------------------------------------
                // FAZ O FRACIONAMENTO DOS PERIODOS 
                // ---------------------------------------------
                string fracao = string.Empty;
                //DEBUG ONLY
                /*
                fracao += $"Total de periodos nesta análise: {seg.Count} {ENTER}";
                fracao += "Passou: ";
                foreach (var item in passou)
                {
                    fracao += $"{item.ToString()}-";

                }
                fracao += ENTERS;
                */
                int fracs = seg.Count; // Numero de periodos 
                if (true) //
                {
     if (tb_empresa.Text != "") fracao += $"Empresa: {tb_empresa.Text}" + ENTER;
     if (fracs > 1) fracao += "Houve necessidade de fracionamento do período pelo perito, tendo em vista os nuances da legislação aplicável" +
                        " ao longo do tempo, conforme se segue: " + ENTER;
                    //else fracao += "Resumo do enquadramento: " + ENTER;
                    // Iterate lista de periodos já fragmentados 
                    for (int i = 0; i < seg.Count; i++)
                    {
                        var item = seg[i];
                        if (fracs > 1) fracao += $"{(i + 1).ToString()}º período: ";
                        if (i == 0) // primeiro periodo 
                            fracao += dbinicio.Value.ToShortDateString(); // data inicial do periodão 
                        else
                            fracao += js[item[0]].inicio.ToShortDateString(); // inicio da janela
                        fracao += " a ";

                        if (i == seg.Count - 1) // ultimo periodo 
                            fracao += dbfim.Value.ToShortDateString(); // data final do periodão 
                        else
                            fracao += js[item[item.Count - 1]].fim.ToShortDateString(); // fim da janela

                        if (passou[js[item[0]].num]) fracao += " - ENQUADRADO";
                        else fracao += " - NÃO ENQUADRADO";
                        fracao += ENTER;
                    }
                    fracao += ENTER;
                }
                tblaudo.Text = fracao + tblaudo.Text;

                // REMOVE ESPAÇOS DUPLOS
                tblaudo.Text = tblaudo.Text.Replace("  ", " ");



                //}
                //try {    }

                //catch (Exception ex)
                //  {
                //    MessageBox.Show("Erro: Parâmetros preenchidos incorretamente / incompletos: " + ex.Message);

                //}

            }
            catch
            {
                fala("Impossível continuar: Dados inseridos são inválidos. Corrigir ou reiniciar o aplicativo.");
            }
	 if (cb_autocopy.Checked) // se estiver marcado, copia o laudo para a área de transferência 
	 {
		Clipboard.SetText(tblaudo.Text);
	 }
	}

        private void button2_Click(object sender, EventArgs e)
        {
   cb_usou_NEN.Checked = false; // reseta checkbox de NEN
	 tbruivalue.Text = "";
	 tb_datas_display.Text = "";
   tb_empresa.Text = "";
            cbvibra.Checked = false;
            gbvinter.Enabled = false;
            gbvibra97.Enabled = false;
            rbmetodo.Checked = true; // reseta metodo do ruido 
            cbbmetodo.SelectedIndex = 4; // reseta combo box metodo ruido
            dbinicio.Value = new DateTime(1992, 01, 01);
            dbfim.Value = new DateTime(2004, 12, 31);
            rtbextrator.Text = "";
            tbouprofctps.Text = tbouprofissaoppp.Text = "";
            tblaudo.Text = "";
            cbruido.Checked = cbeletricidade.Checked = cbusaq.Checked = cbusabio.Checked = 
            cberg2.Checked = cbergo.Checked = cbfrio.Checked = cbfis.Checked = cbrni.Checked = 
            cbusacalor.Checked = cbumidade.Checked =  false;
            while (lbqfinal.Items.Count > 0) apagaum(); // apaga lista de quimicos 
            tab_maintab.SelectedTab = tabPage11;
	 rtbextrator.Focus(); // foca janela do extrator
	}

        private void cbbmetodo_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            cbruneoutro.Checked = true;
        }

        // Combo box que seleciona a metodologia do ruido 
        private void cbbmetodo_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            rbmetodo.Checked = true;
	 if (cbbmetodo.SelectedIndex == 2 && !cb_usou_NEN.Checked)
	 {
		cb_usou_NEN.ForeColor = Color.Red; // muda a cor do texto da checkbox para vermelho
		cb_usou_NEN.Font = new Font(cb_usou_NEN.Font, FontStyle.Bold);
	 }
	 else
	 {
		cb_usou_NEN.ForeColor = SystemColors.ControlText; // reseta para a cor padrão
		cb_usou_NEN.Font = new Font(cb_usou_NEN.Font, FontStyle.Regular);
	 }

	}

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            vequimbox();
        }

        private void rbdaquimi_CheckedChanged(object sender, EventArgs e)
        {
            vequimbox();
        }
        void vequimbox()
        {
            
            if (rbdaquimi.Checked)
            {
                cbquimieve.Checked = false;
                gbnegaquimi.Enabled = false;
                gbdaquimi.Enabled = true;
            }
            else // rbnegaquimi.checked
            {
                
                gbnegaquimi.Enabled = true;
                gbdaquimi.Enabled = false;
            }
            if (cbquimieve.Checked) rbnegaquimi.Checked = true;
        }
        private void rbnegaquimi_CheckedChanged(object sender, EventArgs e)
        {
            vequimbox();
        }

        private void rbeleacima250_CheckedChanged(object sender, EventArgs e)
        {
            if (rbeleacima250.Checked) cbelenaoacima250.Enabled = true;
            else cbelenaoacima250.Enabled = false;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void cbquimico_CheckedChanged(object sender, EventArgs e)
        {
            if (!cbusaq.Checked)
            {
                gbdaquimi.Enabled = false;
                gbnegaquimi.Enabled = false;
                rbnegaquimi.Enabled = false;
                rbdaquimi.Enabled = false;

            }
            else
            {
                rbnegaquimi.Enabled = true;
                rbdaquimi.Enabled = true;
                vequimbox();
            }
        }

        private void cbruido_CheckedChanged(object sender, EventArgs e)
        {
            if (cbruido.Checked)
                gbrmeto.Enabled= gbrunit.Enabled= gbruido.Enabled =cblayout.Enabled = true;
            else
               gbrmeto.Enabled = gbrunit.Enabled = gbruido.Enabled = cblayout.Enabled = false;
   tbruivalue.Focus(); // foca campo de valor do ruido 


	}

        private void cbeletricidade_CheckedChanged(object sender, EventArgs e)
        {
            if (cbeletricidade.Checked) gbele.Enabled = true;
            else gbele.Enabled = false;
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cbusabio_CheckedChanged(object sender, EventArgs e)
        {
            if (cbusabio.Checked) rbbnega.Enabled= gbbprof.Enabled = gbbmotivos.Enabled = 
                 rbbdasaude.Enabled = gbbio3.Enabled = true;
            else gbbprof.Enabled = gbbmotivos.Enabled = gbbio3.Enabled =  false;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void rbbdasaude_CheckedChanged(object sender, EventArgs e)
        {
            if (rbbdasaude.Checked) gbbmotivos.Enabled = false;
            if (!rbbdasaude.Checked) gbbmotivos.Enabled = true;
        }

        private void rbbnega_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
           
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
           
        }

        private void button5_Click(object sender, EventArgs e)
        {
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
           
        }

        private void button7_Click(object sender, EventArgs e)
        {
            tbqo.Text = string.Empty;
           
        }

        private void button9_Click(object sender, EventArgs e)
        {
           
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button10_Click(object sender, EventArgs e)
        {
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            
        }

        private void cbcalor_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void button11_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Até 05/03/97:  Enquadrar acima de 28 graus OU profissões específicas, sempre de fontes " +
                "artificiais. De 06 / 03 / 97 em diante: sempre quantitativo, sendo valor em IBUTG, devendo-se fazer, " +
                "a média ponderada do consumo energético do quadro 3 do anexo da NR-15 conforme a proporção trabalho/repouso " +
                " sendo que em caso de repouso local longe do calor, utiliza-se taxa de metabolismo 100cal/h, e em caso de " +
                "sendo que os limites variam conforme o tipo de atividade(leve, moderada, pesada).");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            MessageBox.Show("I-Trabalhador das Indústrias Metalúrgicas e Mecânicas - aciarias, fundições de ferro e metais não ferrosos,  laminações,  forneiros,  mãos  de  forno,  reservas  de  forno,  fundidores,  soldadores,  lingoneiros,  tenazeiros,  caçambeiros,  amarradores,  dobradores  e  desbastadores;  rebarbadores,  esmerilhadores,  marteleiros  de  rebarbação;  operadores  de  tambores  rotativos  e  outras  máquinas  de rebarbação;\nII-Operadores de máquinas para fabricação de tubos por centrifugação;\nIII-Operadores de pontes rolantes ou de equipamentos para transporte de peças e caçambas com metal liquefeito, nos recintos de aciarias, fundições e laminações; operadores nos fornos de recozimento ou de têmpera-recozedores, temperadores.\nIV-Trabalhadores em ferrarias,  estamparias de  metal a quente e caldeiraria: ferreiros,  marteleiros,  forjadores, estampadores, caldeireiros e prensadores; operadores de forno de recozimento,  de  têmpera,  de  cementação, forneiros, recozedores, temperadores, cementadores; operadores de pontes rolantes ou talha elétrica.\nV-Trabalhadores permanentes relacionados a fabricação de vidros e  cristais: vidreiros  operadores  de  forno,  forneiros,  sopradores, secadores, ou operadores de máquinas de produção de vidros e cristais");
        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void radioButton14_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void rbcacontinuo_CheckedChanged(object sender, EventArgs e)
        {
            cacalcula();
            if (rbcacontinuo.Checked)
            {
                cbbcatrab.Enabled = tbcadescanso.Enabled = lbcatrabalha.Enabled = lbcadescansa.Enabled = false;
                cbbcatrab.SelectedIndex = 0;
            }
            else { cbbcatrab.Enabled = tbcadescanso.Enabled = lbcatrabalha.Enabled = lbcadescansa.Enabled = true; }
        }
        static double catrab = 0; // tempo de trabalho 
        static double cadesc = 0; // tempo de descanso 
        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            cacalcula();
        }
        void cacalcula()
        {
            int modo=1; // descansa em outro local  
            if (rbcanaoinformado.Checked) modo = 1; // descansa no trabalho 
            if (rbcadescansatrabalho.Checked) modo = 2; // descansa no trabalho 
            if (rbcadescansaoutro.Checked) modo = 3; // descansa no trabalho 

            cakcal = cad[cbbcaatividade.SelectedIndex];
            tbcakcal.Text = cakcal.ToString();
            if (cbbcatrab.SelectedIndex == -1) cbbcatrab.SelectedIndex = 0;
            catrab = Convert.ToInt32(cbbcatrab.SelectedItem.ToString());
            cadesc = 60 - catrab;
            tbcadescanso.Text = cadesc.ToString();
            double m = (catrab * cakcal + cadesc * 100) / 60;
            tbcaibutg.Text = Math.Round(m,1).ToString();
            double limite = 0;

            // Se o modo não foi informado, aplica o mais vantajoso 
            if (rbcanaoinformado.Checked) limite = Math.Min(calimite(m, 2), calimite(m, 3));
            else limite = calimite(m, modo);
            tbcalim.Text = limite.ToString();

        }
        private void comboBox2_SelectedIndexChanged_2(object sender, EventArgs e)
        {
            cacalcula();


        }
        // TABELA DE CONSUMO ENERGÉTICO CONFORME ATIVIDADE
        static int cakcal = 0;
        static Dictionary<int, int> cad = new Dictionary<int, int>()
        {
            [0] = 125,
            [1] = 150,
            [2] = 150,
            [3] = 180,
            [4] = 175,
            [5] = 220,
            [6] = 300,
            [7] = 440,
            [8] = 550,
        };
        static double calimite(double kcal,int forma)
        {
            // descansa no local de trabalho
            if (forma == 2)
            {
                if (cakcal >= 440) // trabalho pesado 
                {
                    if (catrab == 60) return 25;
                    if (catrab == 45) return 25.9;
                    if (catrab == 30) return 27.9;
                    if (catrab == 15) return 30;
                }
                else if (cakcal >= 175) // trabalho moderado
                {
                    if (catrab == 60) return 26.7;
                    if (catrab == 45) return 28;
                    if (catrab == 30) return 29.4;
                    if (catrab == 15) return 31.1;
                }
                else // trabalho leve 
                {
                    if (catrab == 60) return 30;
                    if (catrab == 45) return 30.5;
                    if (catrab == 30) return 31.4;
                    if (catrab == 15) return 32.2;
                }
                
            }
            else // Forma == 3, descansa outro local 
            {
                Dictionary<double, double> lim = new Dictionary<double, double>()
                {
                    [0] = 100,
                    [175] = 30.5,
                    [200] = 30,
                    [250] = 28.5,
                    [300] = 27.5,
                    [350] = 26.5,
                    [400] = 26,
                    [450] = 25.5,
                    [500] = 25,
                    [1000] = 25,
                };
                double last = 0;
                foreach (var keypair in lim)
                {
                    if (kcal > last && kcal <= keypair.Key)
                    {
                        return lim[keypair.Key];
                    }
                    last = keypair.Key;
                }
            }
            return -1;
        }
        



        private void cbbcametodo_SelectedIndexChanged(object sender, EventArgs e)
        {
            rbcametodo.Checked = true;
        }

        private void rbcanaoinformado_CheckedChanged(object sender, EventArgs e)
        {
            if (rbcanaoinformado.Checked) { cbbcametodo.SelectedIndex = 3; rbcametodo.Checked = true; }
        }

        private void rbcadescansaoutro_CheckedChanged(object sender, EventArgs e)
        {
            cacalcula();
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("Antes de 1997: Concede para profissões específicas (qualitativo) ou acima de 28 graus (demais) somente de fonte artificial. " +
                "Após 1997: Sempre quantitativo. Trabalho com repouso no próprio local usa limites do quadro nº 1 do anexo 3 da NR-15; " +
                "Trabalho com repouso em local distinto usa limites do quadro nº 2 do anexo 3 da NR-15, fazendo a média ponderada da " +
                "taxa metabólica do quadro nº 3 da NR-15 em relação à quantidade de repouso, considerando 100 kcal/h durante o repouso. " +
                "Caso não exista informação sobre a forma de repouso (no local ou em outro local) será considerado trabalho contínuo e. " +
                "aplicada a tabela mais vantajosa. Não será realizada interpolação nos valores do quadro nº 2 do anexo 3 da NR-15, usa-se " +
                "os limites de IBUTG da coluna da direita para valores menores ou iguais ao metabolismo médio ponderado da coluna da " +
                "esquerda da tabela original.");
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void cbbunit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbbunit.SelectedIndex == 0) { cb_unidade_errada.Checked = false; cb_unidade_errada.Enabled = false; }
            else
            {
                cb_unidade_errada.Enabled = true;
                cb_unidade_errada.Checked = true;
            }

        }

        private void cbunit_CheckedChanged(object sender, EventArgs e)
        {
            if (cbbunit.SelectedIndex == 0) { cb_unidade_errada.Checked = false; cb_unidade_errada.Enabled = false; }
            else cb_unidade_errada.Enabled = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Concede somente até 05/03/1997. Análise sempre qualitativa, sendo necessária exposição permanente, não eventual. " +
                "Praticamente o único tipo de radiação ionizante que gera tempo especial é a luz ultravioleta para atividade " +
                "de soldador. A luz solar produz radiação ultra-violeta entretanto a exposição será considerada intermitente, uma vez que " +
                "mesmo que trabalhe em campo aberto, existem os dias de chuva, tempo nublado, fim de tarde, ou mesmo à noite, quando o trabalhador " +
                "não estará exposto ao sol. Outras formas incomuns de radiação ionizante são: Laser (difícil antes de 1997), radar, ondas de rádio, " +
                "infravermelho, e microondas. Não existe lista de profissões específicas para concessão. Não é necessário uso de nenhuma " +
                "metodologia específica. A partir de 14/10/1996 pode-se exijir metodologia 'inspeção do local de trabalho' e LTCAT/outras " +
                "demonstrações ambientais.");
                
                
                
                
        }

        private void radioButton1_CheckedChanged_1(object sender, EventArgs e)
        {
                    }

        private void groupBox9_Enter(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {

        }

        private void cbrni_CheckedChanged(object sender, EventArgs e)
        {
            if (cbrni.Checked) gbrnitipo.Enabled = gbrnimotivos.Enabled = rbrnida.Enabled = rbrninega.Enabled  =gbrniatividade.Enabled = true;
        }

        private void cbquimicos_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button28_Click(object sender, EventArgs e)
        {
            fala("Mesmo que esta opção esteja marcada, somente será concedido até 05/03/1997.");
        }
        static void fala(string msg)
        {
            MessageBox.Show(msg);
        }

        private void cbbeve_CheckedChanged(object sender, EventArgs e)
        {

        }
        

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
                // MARCA RADIO BUTTON NEGAR 
                foreach (Control c in gbrnimotivos.Controls)
                    if (c is CheckBox)
                    {
                        CheckBox chk = (CheckBox)c;
                        if (chk.Checked)
                            rbrninega.Checked = true;
                    }

            
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            // MARCA RADIO BUTTON NEGAR 
            foreach (Control c in gbrnimotivos.Controls)
                if (c is CheckBox)
                {
                    CheckBox chk = (CheckBox)c;
                    if (chk.Checked)
                        rbrninega.Checked = true;
                }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            // MARCA RADIO BUTTON NEGAR 
            foreach (Control c in gbrnimotivos.Controls)
                if (c is CheckBox)
                {
                    CheckBox chk = (CheckBox)c;
                    if (chk.Checked)
                        rbrninega.Checked = true;
                }
        }

        private void checkBox28_CheckedChanged(object sender, EventArgs e)
        {
            // MARCA RADIO BUTTON NEGAR 
            foreach (Control c in gbrnimotivos.Controls)
                if (c is CheckBox)
                {
                    CheckBox chk = (CheckBox)c;
                    if (chk.Checked)
                        rbrninega.Checked = true;
                }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            // MARCA RADIO BUTTON NEGAR 
            foreach (Control c in gbrnimotivos.Controls)
                if (c is CheckBox)
                {
                    CheckBox chk = (CheckBox)c;
                    if (chk.Checked)
                        rbrninega.Checked = true;
                }
        }

        private void checkBox35_CheckedChanged(object sender, EventArgs e)
        {
            // MARCA RADIO BUTTON NEGAR 
            foreach (Control c in gbrnimotivos.Controls)
                if (c is CheckBox)
                {
                    CheckBox chk = (CheckBox)c;
                    if (chk.Checked)
                        rbrninega.Checked = true;
                }
        }

        private void checkBox29_CheckedChanged(object sender, EventArgs e)
        {
            
            // MARCA RADIO BUTTON NEGAR 
            foreach (Control c in gbrnimotivos.Controls)
                if (c is CheckBox)
                {
                    CheckBox chk = (CheckBox)c;
                    if (chk.Checked)
                        rbrninega.Checked = true;
                }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            fala("O texto será gerado na caixa de texto dos laudos, apagando seu conteúdo.");
        }

        private void button31_Click(object sender, EventArgs e)
        {
            tblaudo.Clear();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            tblaudo.Text = "A tarefa atual inserida no no PMF-Tarefas contém intervalo de tempo que abrange mais de um período de exposição a agentes nocivos do campo 14.1" +
                " do PPP, em desacordo com o item 3.12 da cartilha de aposentadoria especial e ofício circular conjunto nº 44/DIRBEN/DIRAT/INSS" +
                " de 01/11/2019, sendo necessária a subdivisão do intervalo de tempo em mais de uma tarefa, de forma que cada tarefa contemple apenas " +
                "um período. ";
                
        }

        private void button23_Click(object sender, EventArgs e)
        {
            fala("Utilize esta opção quando o PPP incluir agentes que não são enquadráveis, como ergonômico, físico, acidentes, etc.");
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            fala("O anexo do Decreto 53831/64 cita especificamente a atividade de soldador com arco elétrico e com oxiacetilenio como exposição a radiações não ionizantes. A exposição nesse " +
                "caso será a ondas ultra-violeta.");
        }

        private void button20_Click(object sender, EventArgs e)
        {
            fala("O anexo do Decreto 53831/64 cita especificamente a atividade de aeroviários de manutenção de aeronaves e motores, turbo-hélices e outros. A exposição nesse caso"+
                " será a ondas de radar.");
        }

        private void rbrnisoldador_CheckedChanged(object sender, EventArgs e)
        {
            if (rbrnisoldador.Checked) cbrniuv.Checked = true;
        }

        private void rbrniaeroviario_CheckedChanged(object sender, EventArgs e)
        {
            if (rbrniaeroviario.Checked) cbrniradar.Checked = true;
        }

        private void cbrniagentenenhum_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void button33_Click(object sender, EventArgs e)
        {
            fala("Use esta opção se não faz nenhum sentido o PPP colocar radiações não ionizantes como agente nocivo.");
        }

        private void button32_Click(object sender, EventArgs e)
        {
            fala("use esta opção para citar outros tipos de radiações não ionizantes, como ondas de TV, ondas de celular, ultra-som, " +
                "luz negra, etc. Caso não seja especificada qual, será gerado texto genérico.");
        }

        private void button15_Click(object sender, EventArgs e)
        {
            fala("Fontes: Luz solar, solda elétrica ou de axiacetileno, fotopolimerizador de resinas (dentistas, indústria química), luz negra (casas noturnas), etc.");
        }

        private void button16_Click(object sender, EventArgs e)
        {
            fala("Fontes comuns: Emissores de rádio, TV, celular, radioamador, microfones sem fio, telefone sem fio, celular, WI-FI, walkie-talkie, rádio de polícia ou ambulância. Antena parabólica não emite ondas.");
        }

        private void button17_Click(object sender, EventArgs e)
        {
            fala("Praticamente só existente em aeroportos e instalações militares, mas também em radares de velocidade rodoviária antigos (os novos usam laser).");
        }

        private void button18_Click(object sender, EventArgs e)
        {
            fala("Forno");
        }

        private void button19_Click(object sender, EventArgs e)
        {
            fala("Não confundir com magnetismo. Todas as radiações não ionizantes são ondas eletromagnéticas.");
        }

        private void button21_Click(object sender, EventArgs e)
        {
            fala("Presente em 'radares' portáteis de velocidade, dispositivos de medição na indústria, máquinas de corte e gravação.");
        }

        private void button24_Click(object sender, EventArgs e)
        {
            fala("Marque esta caixa se a exposição foi a outro tipo de ondas, como Raio-X (ionizante), ultra-som, sonar, etc.");
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void tbrniagenteoutros_TextChanged(object sender, EventArgs e)
        {
            if (tbrniagenteoutros.Text != string.Empty) cbrniagenteoutro.Checked = true;
        }

        private void button18_Click_1(object sender, EventArgs e)
        {
            tblaudo.Text = $"{dbinicio.Value.ToShortDateString()} a {dbfim.Value.ToShortDateString()}: Conclusão: Período " +
                $"NÃO ENQUADRADO. Elementos para não enquadramento: O PPP não contém informação sobre agentes nocivos nos campos 15.2 a 15.5, " +
                $"aonde deveria constar a descrição dos agentes nocivos ao qual o trabalhador foi exposto, " +
                $"impossibilitando a identificação inequívoca de que fatores, e em qual intensidade ocorreram durante a jornada de " +
                $"trabalho. ";
 




        }

        private void tblaudo_TextChanged(object sender, EventArgs e)
        {
            int charcount = tblaudo.Text.Length; // conta caracteres 
            tbchars.Text = $"{charcount.ToString()}/3990"; // escreve na text box 
            if (charcount > 3990)
            {
                tbchars.ForeColor = Color.Red; // muda para vermelho se passar 
                
            }
            else tbchars.ForeColor = Color.Black;
        }

        private void button34_Click(object sender, EventArgs e)
        {
            var s = tblaudo.Text;
            s = s.Replace("1997", "97");
            s = s.Replace("1998", "98");
            s = s.Replace("1999", "99");
            s = s.Replace("1996", "96");
            s = s.Replace("1995", "95");
            s = s.Replace("1994", "94");
            s = s.Replace("1993", "93");
            s = s.Replace("1992", "92");
            s = s.Replace("1991", "91");
            s = s.Replace("1964", "64");
            s = s.Replace("2003", "03");
            s = s.Replace("2004", "04");
            s = s.Replace("1978", "78");
            s = s.Replace("utilizando-se", "usando");
            s = s.Replace("portaria", "port");
            s = s.Replace("Decreto", "Dec");
            s = s.Replace("IN 77 / 2015 / INSS", "IN 77 ");
            s = s.Replace(" da FUNDACENTRO", "");
            s = s.Replace(", ", " ");
            s = s.Replace("5 de março de 1997", "5/03/97");
            s = s.Replace(" do MTE", "");
            s = s.Replace("/2015/INSS", "");
            s = s.Replace("Normativa vigente", "Fundamentação");
            s = s.Replace("entretanto", "mas");
            s = s.Replace("utilizada", "usada");
            s = s.Replace("cód. ", "");
            s = s.Replace("Elementos", "Fatores");
            s = s.Replace("nº", "");
            s = s.Replace("Art.", "Art");
            s = s.Replace(" volts", "v");
            s = s.Replace("\n\n", "\n");
            s = s.Replace(" ocupacional", "");
            s = s.Replace("- Período", "-");
            s = s.Replace(":  Período", ":");
            s = s.Replace("  ", " ");
            //s = s.Replace("NR-15", "NR15");
            tblaudo.Text = s;
        }

        private void button33_Click_1(object sender, EventArgs e)
        { // botão metodologia frio
            fala("Após 06/03/1997: Não enquadra nunca.\n" +
                "Antes de 06/03/1997: Qualitativo para trabalho em câmara fria, câmara frigorífica, ou fabricação de gelo.\n" +
                "Antes de 14/10/1996: Além do qualitativo (profissões acima) também permite quantitativo abaixo de 12 graus, sem exigir " +
                "metodologia, somente fontes artificiais.");
        }

        private void button35_Click(object sender, EventArgs e)
        {
            fala("É comum o PPP incluir engenheiros, inspetores, encarregados, supervisores e outras funções que não ficam" +
                " realmente expostas ao frio e nem sequer entram na câmara fria.");
        }

        private void button36_Click(object sender, EventArgs e)
        {
            fala("É comum o PPP incluir gerentes, administrativos, diretores e chefias.");
        }

        private void button37_Click(object sender, EventArgs e)
        {
            fala("Marque somente uma opção. Caso seja marcado, a temperatura informada e a profissão serão desconsideradas.");
        }

        private void button38_Click(object sender, EventArgs e)
        {
            fala("Marque esta opção se o trabalhador ficava exposto ao frio somente durante alguns períodos da jornada.");
        }

        private void button25_Click(object sender, EventArgs e)
        {
            fala("Não houve exposição a radiação não ionizante, o PPP não faz nenhum sentido. ");
        }

        private void button27_Click(object sender, EventArgs e)
        {
            fala("Ficou exposto somente durante alguns períodos da jornada, intercalando dias ou horários exercendo atividade exposta e outras" +
                "atividades administrativas, não operacionais, ou longe das radiações. Um exemplo é um mecânico de manutenção que efetuava soldas"+
                " eventualmente.");
        }

        private void button26_Click(object sender, EventArgs e)
        {
            fala("Exposição somente a raios solares não devem ser enquadradas, porque existem dias de chuva, nublados, crepúsculo, em que não há " +
                "exposição às radiações ultra-violetas, de forma que a exposição fica intermitente.");
        }

        private void cbfrio_CheckedChanged(object sender, EventArgs e)
        {
            if (cbfrio.Checked)
            {
                gbfrio1.Enabled = gbfrio2.Enabled = gbfrio3.Enabled = true;
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            fala("O texto fica um pouco diferente para profissões da área de saúde (cita contato com pacientes) e para áreas distintas.");
        }

        private void button40_Click(object sender, EventArgs e)
        {
            fala("Os códigos são diferentes para patógenos de origem humana ou animal até 05/03/97.");
        }

        private void radioButton1_CheckedChanged_2(object sender, EventArgs e)
        {
            if (rbbveterinario.Checked) rbbanimal.Checked = true;
            
        }

        private void button41_Click(object sender, EventArgs e)
        {
            fala("Se não for marcado nenhum, e a opção 'negar' estiver marcada, vai gerar um texto genérico.");
        }

        private void button47_Click(object sender, EventArgs e)
        {
            fala("ex: clima úmido, água do mar, água da chuva, trabalho em rios, banhados e mangues.");
        }

        private void button46_Click(object sender, EventArgs e)
        {
            fala("Mesmo que assinalado, só vai indeferir entre 14/10/96 a 5/3/97, quando era obrigatório.");
        }

        private void cbfriosemfrio_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button45_Click(object sender, EventArgs e)
        {
            fala("É comum o PPP incluir engenheiros, inspetores, encarregados, supervisores e outras funções que não ficam" +
               " realmente expostas à umidade e nem sequer entram em contato com a água, em empresas em que os trabalhadores da" +
               "ponta ficam expostos ao agente nocivo.");
        }

        private void button44_Click(object sender, EventArgs e)
        {
            fala("É comum o PPP incluir gerentes, administrativos, diretores e chefias de empresas aonde os funcionários ficam " +
                "expostos aos agentes de risco.");
        }

        private void button42_Click(object sender, EventArgs e)
        {
            fala("Marque esta opção se o trabalhador ficava exposto à umidade somente durante alguns dias ou períodos da jornada.");
        }

        private void button43_Click(object sender, EventArgs e)
        {
            fala("Se não for marcado nenhuma opão abaixo, vai enquadrar tudo até 05/03/97.");
        }

        private void button49_Click(object sender, EventArgs e)
        {
            fala("Mesmo que assinalado, vai verificar a existência de LTCAT e demonstrações ambientais (última checkbox)");
        }

        private void button43_Click_1(object sender, EventArgs e)
        {
            fala("O enquadramento para umidade ocorrerá em caso de trabalhos em contato direto e permanente com água, sempre " +
                "de fonte artificial. A avaliação será qualitativa e somente até 05/03/97. Após este período nunca enquadra." +
                "Durante um curto período entre 14/10/96 a 5/3/97 era exigido LTCAT ou outras demonstrações ambientais.");
        }

        private void button50_Click(object sender, EventArgs e)
        {
            fala("Mesmo que ultrapasse o espaço não tem problema. O aplicativo fará uma introdução, e depois anexa o motivo escrito na caixa de texto.");
        }

        private void cbumidade_CheckedChanged(object sender, EventArgs e)
        {
            if (cbumidade.Checked) gbumi.Enabled = true;
        }

        private void button51_Click(object sender, EventArgs e)
        {
            fala("Selecionar qualquer metodologia diferente da NR-15 ou NHO-01 sempre irá indeferir o enquadramento. Algumas vezes o PPP " +
                "cita a metodologia como 'decibelímetro' ou 'dosimetria' mas na LTCAT ou demais demonstrações ambientais é demonstrado que " +
                "foi na realidade utilizada a técnica correta, NR-15 ou NHO-01, portanto é necessário verificar com atenção.");
        }

        private void button52_Click(object sender, EventArgs e)
        {
            fala("Será sempre considerado o menor valor para fins de enquadramento. Existem duas correntes quanto a valores de máximo e mínimo: " +
                "os que não aceitam este tipo de informação, e os que utilizam o menor valor. ");
        }

        private void button53_Click(object sender, EventArgs e)
        {
            fala("O aplicativo utiliza o quadro 17, da página 91 do manual de aposentadoria especial do INSS/2018. Os limites de tolerância" +
                " variam ao longo dos anos, indo de 80 a 90dB(A) conforme a época. A técnica NR-15 era obrigatória até 18/11/03, e a NHO-01 " +
                "passou a ser obrigatória a partir de 01/01/04. Houve um período de transição de cerca de um mês entre 19/11/03 a 31/12/03, quando" +
                " foram aceitas as duas metodologias. Utilizar a metodologia NR-15 após 01/01/04 nunca será permitido; mas usar a metodologia " +
                "NHO-01 antes de 19/11/03 é aceitável desde que exista a informação expressa de que não houve alteração nas rotinas de trabalho, " +
                "maquinário ou layout da empresa (laudo extemporâneo).");
        }

        private void button54_Click(object sender, EventArgs e)
        {
            fala("Marcando esta opção será aceita a metodologia NHO-01 para qualquer período, mesmo anteriores a 19/11/03. Usar a metodologia " +
                "NHO-01 antes de 19/11/03 é aceitável desde que exista a informação expressa de que não houve alteração nas rotinas de trabalho, " +
                "maquinário ou layout da empresa (laudo extemporâneo). A FUNDACENTRO somente publicou a norma NHO-01 no ano de 2001, portanto " +
                "informação de uso desta normativa anteriormente a 2001 poderia também ser indício de um PPP gracioso. ");  
        }

        private void button55_Click(object sender, EventArgs e)
        {
            fala("Use esta opção se o PPP não informar o valor do ruído. Irá indeferir, uma vez que o enquadramento só pode " +
                "ser quantitativo.");
        }

        private void tbrniprof_TextChanged(object sender, EventArgs e)
        {
            if (tbrniprof.Text != string.Empty) rbrniprofoutra.Checked = true;
        }

        private void tbrninegaoutro_TextChanged(object sender, EventArgs e)
        {
            if (tbrninegaoutro.Text != string.Empty) cbrninegaoutro.Checked = true;
        }

        private void tbruimetodooutro_TextChanged(object sender, EventArgs e)
        {
            if (tbruimetodooutro.Text != string.Empty) rbruimetodooutro.Checked = true;
        }

        private void tbmin_TextChanged(object sender, EventArgs e)
        {
            rbruiminmax.Checked = true;
        }

        private void tbruivalue_TextChanged(object sender, EventArgs e)
        {
            tbruiunico.Checked = true;
        }

        private void tbeleunico_TextChanged(object sender, EventArgs e)
        {
            rbeleunico.Checked = true;
        }

        private void tbelemin_TextChanged(object sender, EventArgs e)
        {
            rbeleminmax.Checked = true;
        }

        private void tbeleoutro_TextChanged(object sender, EventArgs e)
        {
            cbeleoutro.Checked = true;
        }

        private void tbqotg_TextChanged(object sender, EventArgs e)
        {
            cbqotg.Checked = true;
        }

        private void tbqpnn_TextChanged(object sender, EventArgs e)
        {
            cbqppn.Checked = true;
        }

        private void tbqqal_TextChanged(object sender, EventArgs e)
        {
            cbqqal.Checked = true;
        }

        private void tbqo_TextChanged(object sender, EventArgs e)
        {
            cbqo.Checked = true;
        }

        private void button56_Click(object sender, EventArgs e)
        {
            fala("Esta aba de químicos é apenas um esboço, uma referência para gerar alguns textos que podem ser utilizados como " +
                "arcabouço para uma análise manual propriamente dita. Análise de agentes químicos é extremamente complexa e ainda está em"+
                " fase de desenvolvimento.");
        }

        private void tbboutra_TextChanged(object sender, EventArgs e)
        {
            rbboutra.Checked = true;
        }

        private void tbcaunico_TextChanged(object sender, EventArgs e)
        {
            rbcavalorunico.Checked = true;
        }

        private void tbcamin_TextChanged(object sender, EventArgs e)
        {
            rbcaminmax.Checked = true;
        }

        private void tbfriunico_TextChanged(object sender, EventArgs e)
        {
            rbfrivalorunico.Checked = true;
        }

        private void tbfrimin_TextChanged(object sender, EventArgs e)
        {
            rbfriminmax.Checked = true;
        }

        private void tbumioutro_TextChanged(object sender, EventArgs e)
        {
            rbumioutro.Checked = true;
        }

        private void tbergo2_TextChanged(object sender, EventArgs e)
        {
            cberg2.Checked = true;
        }

        private void button57_Click(object sender, EventArgs e)
        {
            printlb(lbqoit);
        }
        public void wait(int milliseconds) // It will wait number of miliseconds. 
        {
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds <= milliseconds)
            {
                Application.DoEvents();
            }
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            filtra(tbqbusca.Text, qaIVqual, ref lbqanexoivqual);
            filtra(tbqbusca.Text, qa64, ref lbq64);
            filtra(tbqbusca.Text, qoit, ref lbqoit);
            filtra(tbqbusca.Text, qlinach, ref lbqlinach);
            filtra(tbqbusca.Text, qaivquant, ref lbqivquant);
            
        }
        // PRINT A LISTBOX TO TBLAUDO.TEXT
        void printlb(ListBox lb,int delay = 10)
        {
            //tblaudo.Text = "";
            for (int i = 0; i < lb.Items.Count; i++)
            {
                tblaudo.Text += lb.Items[i].ToString() + ENTER;
                wait(delay);
            }
        }
        private void button58_Click(object sender, EventArgs e)
        {
            printlb(lbq64);
            
        }

        private void button59_Click(object sender, EventArgs e)
        {
            printlb(lbqlinach);
        }

        private void cbusaq_CheckedChanged(object sender, EventArgs e)
        {
            gbnegaquimi.Enabled = gbdaquimi.Enabled = rbnegaquimi.Enabled = rbdaquimi.Enabled = true;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            printlb(lbqanexoivqual);
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            printlb(lbqanexoivqual);
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            printlb(lbqlinach);
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            printlb(lbq64);
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            printlb(lbqivquant);
        }

        private void button56_Click_1(object sender, EventArgs e)
        {
            pinta(0);
        }

        private void checkBox6_CheckedChanged_1(object sender, EventArgs e)
        {// decreto 64
            if (cbq64.Checked) qagflag[0, 0] = true;
            else qagflag[0, 0] = false;
            pinta(0);
        }

        private void checkBox11_CheckedChanged_1(object sender, EventArgs e)
        { // oit
            if (cbqoit.Checked) qagflag[0, 1] = true;
            else qagflag[0, 1] = false;
            pinta(0);
        }

        private void lbqanexoivqual_SelectedIndexChanged(object sender, EventArgs e)
        {
            tbqagente.Text = lbqanexoivqual.SelectedItem.ToString();
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        { // 
            if (cbq97q.Checked) qagflag[0, 2] = true;
            else qagflag[0, 2] = false;
            pinta(0);
        }

        private void checkBox5_CheckedChanged_1(object sender, EventArgs e)
        {
            if (cbqquantok.Checked) qagflag[0, 3] = true;
            else qagflag[0, 3] = false;
            pinta(0);
        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            if (cbqquantacima.Checked) qagflag[0, 4] = true;
            else qagflag[0, 4] = false;
            pinta(0);
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (cbqlinach.Checked) { qagflag[0, 5] = true; cbq97q.Checked = true; } // quali anterior tambem 
            else qagflag[0, 5] = false;
            pinta(0);
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        { // nao permanente
            if (cbqnaoperma.Checked) qagflag[0, 6] = true;
            else qagflag[0, 6] = false;
            pinta(0);
        }

        private void checkBox12_CheckedChanged_1(object sender, EventArgs e)
        {// generico 
            if (cbq2gen.Checked) qagflag[0, 7] = true;
            else qagflag[0, 7] = false;
            pinta(0);
        }

        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {// aceita metodologia 
            if (cbqaceitametodo.Checked) qagflag[0, 8] = true;
            else qagflag[0, 8] = false;
            pinta(0);
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {// nr 15 
            if (checkBox16.Checked) qagflag[0, 9] = true;
            else qagflag[0, 9] = false;
            pinta(0);
        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {// NHO QIMICO 
            if (checkBox17.Checked) qagflag[0, 10] = true;
            else qagflag[0, 10] = false;
            pinta(0);
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            limpaquim();

        }
        void limpaquim()
        {
            for (int i = 0; i < 20; i++) qagflag[0, i] = false;
            tbqagente.Text = tbqbusca.Text = string.Empty;
            cbq64.Checked = cbq97q.Checked = cbqlinach.Checked = cbqoit.Checked =
            cbqaceitametodo.Checked = cbqquantacima.Checked = cbqquantok.Checked = false;
            cbqnaoperma.Checked = cbq2gen.Checked = false;
            pinta(0);
        }

        private void lbq64_SelectedIndexChanged(object sender, EventArgs e)
        {
            tbqagente.Text = lbq64.SelectedItem.ToString();

        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            if (tbqagente.Text != string.Empty)
            {
                lbqfinal.Items.Add(tbqagente.Text);
                for (int i = 0; i < 20; i++)
                    qagflag[lbqfinal.Items.Count, i] = qagflag[0, i];
                cbusaq.Checked = true;
            }
            //fala($"{tbqagente.Text} será movido para a lista da aba anterior.");
            limpaquim();
            
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            apagaum();
            
        }
        void apagaum()
        {

            int sumiu = lbqfinal.SelectedIndex;
            if (sumiu == -1 && lbqfinal.Items.Count > 0) sumiu = lbqfinal.Items.Count - 1;
            if (sumiu != -1) // se tiver algo selecioinado 
            {
                lbqfinal.Items.RemoveAt(sumiu);

                for (int i = sumiu + 1; i < 20; i++)
                {
                    for (int j = 0; j < 20; j++)
                        qagflag[i, j] = qagflag[i + 1, j];
                }

            }
        }

        private void lbqoit_SelectedIndexChanged(object sender, EventArgs e)
        {
            tbqagente.Text = lbqoit.SelectedItem.ToString();
        }

        private void lbqlinach_SelectedIndexChanged(object sender, EventArgs e)
        {
            tbqagente.Text = lbqlinach.SelectedItem.ToString();
        }

        private void lbqivquant_SelectedIndexChanged(object sender, EventArgs e)
        {
            string produto = lbqivquant.SelectedItem.ToString();
            tbqagente.Text = produto;
            tbqlim.Text = lim[produto];
        }

        private void dbfim_ValueChanged(object sender, EventArgs e)
        {
            
            check97();
        }

        void check97()
        {
            pinta(0);
            // VERIFICA SE INICIA APÓS 1997
            var novadata = new DateTime(1997, 03, 05); // expire date 
            if (DateTime.Compare(dbinicio.Value, novadata) > 0)
            { // inicia após 97 
                gbca97.Enabled = false;
            }
            else
            { // inicia antes de 97 

                gbca97.Enabled = true;
            }
            if (DateTime.Compare(dbfim.Value, novadata) > 0)
            { // termina antes de 97
                gbcalor98a.Enabled = gbcalor98b.Enabled = true;

            }
            else
            { // termina após 97
                gbcalor98a.Enabled = gbcalor98b.Enabled = false;

            }
        }
        private void dbinicio_ValueChanged(object sender, EventArgs e)
        {
            check97();

        }

        private void button56_Click_2(object sender, EventArgs e)
        {
            fala("1. Procure o elemento químico utilizando a caixa de busca.\n" +
                "2. Escreva manualmente ou clique numa das listas para copiar o nome.\n" +
                "3. Marque as caixas 'checkbox' de todas as listas em que encontrar o elemento. \n" +
                "4. Caso o elemento esteja na lista anexo-IV / quantitativo, marque se acima do limite. \n" +
                "   Os períodos enquadrados serão indicados na cor verde.\n" +
                "5. Marque as caixas de motivos adicionais de indeferimento (não permanente, termo genérico...)\n" +
                "6. Aperte o botão 'Adiciona' para finalizar este agente químico, e transferir o elemento para a lista" +
                " da aba anterior. Poderão ser adicionados outros elementos se necessário.");
        }

        private void lbqfinal_SelectedIndexChanged(object sender, EventArgs e)
        {
            pinta(lbqfinal.SelectedIndex + 1, 2);
        }

        private void button58_Click_1(object sender, EventArgs e)
        { // apaga todos da lista final do quimico 
            while (lbqfinal.Items.Count > 0) apagaum();
        }

        private void button59_Click_1(object sender, EventArgs e)
        {

            tblaudo.Text = $"{dbinicio.Value.ToShortDateString()} a {dbfim.Value.ToShortDateString()}: Conclusão: Período NÃO ENQUADRADO. " +
                $"Elementos para não enquadramento: Não foi apresentado o PPP " +
                "para o período citado. Fundamentação legal: O Decreto nº 4.032, de 26 de novembro de" +
                $" 2001 determinou que a comprovação da efetiva exposição do segurado aos agentes nocivos fosse feita mediante formulário" +
                $" denominado Perfil Profissiográfico Previdenciário – PPP, na forma estabelecida pelo INSS, que estabeleceu este " +
                $"formulário por meio da Instrução Normativa nº 99 / INSS / DC, de 5 de dezembro de 2003, que passou a vigorar a partir" +
                $" de 1º de janeiro de 2004, aonde consta que a efetiva exposição aos agentes nocivos deve ser informada nos campos " +
                $"15.2 do PPP (tipo de agente), 15.3 (fator de risco), 15.4 (intensidade / concentração), e finalmente 15.5 (técnica utilizada).";
        }

        private void button61_Click(object sender, EventArgs e)
        {
            fala("As opções abaixo irão gerar indeferimento independente do que for marcado acima.");
        }

        private void groupBox5_Enter_1(object sender, EventArgs e)
        {

        }

        private void checkBox12_CheckedChanged_2(object sender, EventArgs e)
        {
            var novadata = new DateTime(2014, 08, 13); // expire date 
            if (cbvibra.Checked)
            {
                if (DateTime.Compare(dbfim.Value, novadata) > 0)
                { // inicia após 97 
                    fala("Agente disponível somente até 13/08/2014.");
                    cbvibra.Checked = false;
                }
                else
                {
                    gbvibra97.Enabled = true;
                    gbvinter.Enabled = true;
                }
            }
            else
            {
                gbvibra97.Enabled = false;
                gbvinter.Enabled = false;
            }
        }

        private void button60_Click(object sender, EventArgs e)
        {
            fala("Insira na caixa abaixo texto contendo a data de início e fim do periodo. O extrator irá " +
                "tentar localizar as datas no texto e utilizá-las. Clique o botão 'OK' para ver um exemplo.");
            rtbextrator.Text = $"NOME DA EMPRESA (INSS/GET){ENTER}MIROSOFT ANÁLISE E DESENVOLVIMENTO DE SISTEMAS LTDA{ENTERS}" +
                $"DATA DE INICIO DO PERÍODO (INSS/GET){ENTER}01 / 07 / 1989{ENTERS}DATA FIM DO PERÍODO (INSS/GET){ENTER}" +
                $"10 / 10 / 1991{ENTER}";

        }
	// EXTRATOR DE DATA ATUALIZADO PARA NOVO FORMATO INSS
	private void richTextBox2_TextChanged(object sender, EventArgs e)
	{
	 try
	 {
		string texto = rtbextrator.Text;

		// Limpa variáveis
		string nomeEmpresa = "";
		string dataInicio = "";
		string dataFim = "";

		// Divide o texto em linhas para facilitar a análise
		string[] linhas = texto.Split(new string[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries);

		for (int i = 0; i < linhas.Length; i++)
		{
		 string linha = linhas[i].Trim();

		 // Procura pelo nome da empresa
		 if (linha.Contains("NOME DA EMPRESA") && linha.Contains("INSS"))
		 {
			// A próxima linha deve conter o nome da empresa
			if (i + 1 < linhas.Length)
			{
			 nomeEmpresa = linhas[i + 1].Trim();
			}
		 }

		 // Procura pela data de início
		 else if (linha.Contains("DATA DE INICIO DO PERÍODO") && linha.Contains("INSS"))
		 {
			// A próxima linha deve conter a data de início
			if (i + 1 < linhas.Length)
			{
			 dataInicio = linhas[i + 1].Trim();
			}
		 }

		 // Procura pela data de fim
		 else if (linha.Contains("DATA FIM DO PERÍODO") && linha.Contains("INSS"))
		 {
			// A próxima linha deve conter a data de fim
			if (i + 1 < linhas.Length)
			{
			 dataFim = linhas[i + 1].Trim();
			}
		 }
		}

		// Método alternativo: busca diretamente no texto usando regex-like approach
		if (string.IsNullOrEmpty(dataInicio) || string.IsNullOrEmpty(dataFim))
		{
		 // Busca padrões de data no formato dd/mm/yyyy
		 var datasEncontradas = new List<string>();

		 foreach (string linha in linhas)
		 {
			// Verifica se a linha contém uma data no formato dd/mm/yyyy
			if (System.Text.RegularExpressions.Regex.IsMatch(linha.Trim(), @"^\d{2}/\d{2}/\d{4}$"))
			{
			 datasEncontradas.Add(linha.Trim());
			}
		 }

		 // Se encontrou exatamente 2 datas, assume que são início e fim
		 if (datasEncontradas.Count >= 2)
		 {
			dataInicio = datasEncontradas[0];
			dataFim = datasEncontradas[1];
		 }
		}

		// Variáveis para controlar sucesso da extração
		bool sucessoInicio = false;
		bool sucessoFim = false;

		// Tenta converter e popular os campos
		if (!string.IsNullOrEmpty(dataInicio))
		{
		 if (DateTime.TryParse(dataInicio, out DateTime dtInicio))
		 {
			dbinicio.Value = dtInicio;
			sucessoInicio = true;
		 }
		}

		if (!string.IsNullOrEmpty(dataFim))
		{
		 if (DateTime.TryParse(dataFim, out DateTime dtFim))
		 {
			dbfim.Value = dtFim;
			sucessoFim = true;
		 }
		}

		// Altera a cor dos labels baseado no sucesso da extração
		if (sucessoInicio)
		{
		 label2.ForeColor = Color.Green;  // Azul para sucesso
		}
		else
		{
		 label2.ForeColor = Color.Red;   // Vermelho para erro
		}

		if (sucessoFim)
		{
		 label3.ForeColor = Color.Green;  // Azul para sucesso
		}
		else
		{
		 label3.ForeColor = Color.Red;   // Vermelho para erro
		}

		// Se ambas datas foram extraídas com sucesso, exibe no formato desejado
		if (sucessoInicio && sucessoFim)
		{
		 tb_datas_display.Text = $"{dbinicio.Value:dd/MM/yyyy} a {dbfim.Value:dd/MM/yyyy}";
		}


		// Popular o nome da empresa na textbox tb_empresa
		if (!string.IsNullOrEmpty(nomeEmpresa) && tb_empresa != null)
		{
		 tb_empresa.Text = nomeEmpresa;
		}
	 }
	 catch (Exception ex)
	 {
		// Em caso de erro, não faz nada (silent fail como no código original)
		// Pode adicionar um log ou debug se necessário
	 }
	}

	private void button69_Click(object sender, EventArgs e)
	{
   fala("Períodos e Metodologias para Ruído:\r\nAté 18/11/2003:\r\n\r\nMetodologia: NR-15 (obrigatória)\r\nNão exigia NEN\r\n\r\n19/11/2003 a 31/12/2003:\r\n\r\nPeríodo de transição - eram aceitas duas metodologias:\r\n\r\nNR-15 (tradicional) OU\r\nNHO-01 da FUNDACENTRO (facultativa)\r\n\r\n\r\nSe usasse NHO-01, deveria usar NEN\r\n\r\nA partir de 1º/1/2004:\r\n\r\nMetodologia NHO-01 tornou-se obrigatória\r\nNEN passou a ser obrigatório\r\nLimite de tolerância: NEN superior a 85 dB(A)\r\n\r\nPontos importantes do manual:\r\n\r\n\"Para períodos a serem analisados a partir de 1º/1/2004 é obrigatório que seja utilizada a metodologia da NHO-01 da FUNDACENTRO, devendo estar consignado no PPP os valores de NPS expressos em NEN.\"\r\nA menção do uso da NEN deve constar no campo 15.4 (intensidade/concentração) ou no campo 15.5 (técnica utilizada) do PPP\r\nA simples menção da NHO-01 sem especificar o uso do NEN não é aceita, pois a NHO-01 inclui outras formas de aferição como Leq e TWA");
	}

	private void button70_Click(object sender, EventArgs e)
	{


	 MessageBox.Show("Este programa está sendo disponibilizado gratuitamente e foi fruto de meses de programação, testes e refinamento.\n\n" +
		"Além do tempo em si, houve também custos com hospedagem do servidor e assinaturas dos serviços utilizados no projeto.\n\n" +
		"Se quiser demonstrar sua gratidão me pagando uma pizza, fico feliz! Você pode fazer um Pix de qualquer valor para o \nCPF 021.864.729-88.",
		 "Agradecimento",
		 MessageBoxButtons.OK,
		 MessageBoxIcon.Information
	 );
	}

	// Método de criptografia por substituição simples de caracteres + reverse.
	// Usado para proteger strings sensíveis no código-fonte contra edições ou leitura direta.
	// Marcador #CM# (Cifra Montanha) indica texto cifrado.
	// Adiciona REVERSE para deixar mais divertido e seguro!
	// ----------------------------------------------------
	// MÉTODO CIFRA MONTANHA - SUBSTITUIÇÃO SIMPLES + REVERSE
	// ----------------------------------------------------
	public static string cifraMontanha(string texto)
	{
	 // Gabarito de substituição
	 string original = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,!?;. !@#$%¨&*()-=][:<>";
	 string cifrado = "WLygjSDROVAvbExsh;#itkMYcXwNU!Tzploa,CZIndK?JPBeQGfuFHmr$)(<¨%:@*!^q&-].=][>";

	 // Verifica se o texto começa com o marcador #CM#
	 bool ehTextoCifrado = texto.StartsWith("#CM#");

	 string textoParaProcessar;
	 bool cifrando;

	 if (ehTextoCifrado)
	 {
		// É texto cifrado - remover marcador e descriptografar
		textoParaProcessar = texto.Substring(4); // Remove "#CM#"
		cifrando = false;
	 }
	 else
	 {
		// É texto normal - cifrar
		textoParaProcessar = texto;
		cifrando = true;
	 }

	 StringBuilder resultado = new StringBuilder();

	 // Se está cifrando, primeiro aplica o reverse
	 if (cifrando)
	 {
		char[] chars = textoParaProcessar.ToCharArray();
		Array.Reverse(chars);
		textoParaProcessar = new string(chars);
	 }

	 // Processa cada caractere fazendo a substituição
	 foreach (char c in textoParaProcessar)
	 {
		if (cifrando)
		{
		 // Cifrando: procura na string original e substitui pela cifrada
		 int index = original.IndexOf(c);
		 if (index >= 0)
		 {
			resultado.Append(cifrado[index]);
		 }
		 else
		 {
			// Se não encontrou, mantém o caractere original
			resultado.Append(c);
		 }
		}
		else
		{
		 // Decifrando: procura na string cifrada e substitui pela original
		 int index = cifrado.IndexOf(c);
		 if (index >= 0)
		 {
			resultado.Append(original[index]);
		 }
		 else
		 {
			// Se não encontrou, mantém o caractere original
			resultado.Append(c);
		 }
		}
	 }

	 string textoFinal = resultado.ToString();

	 if (cifrando)
	 {
		// Se estava cifrando, adiciona o marcador no início
		return "#CM#" + textoFinal;
	 }
	 else
	 {
		// Se estava decifrando, aplica reverse para voltar ao original
		char[] chars = textoFinal.ToCharArray();
		Array.Reverse(chars);
		return new string(chars);
	 }
	}

	private void button71_Click(object sender, EventArgs e)
	{
    tblaudo.Text = cifraMontanha(tblaudo.Text);

	}

	private void button72_Click(object sender, EventArgs e)
	{
   fala("Se marcada, o conteúdo do laudo será copiado automaticamente para a área de transferência ao finalizar o preenchimento");
	}

	// Exemplos de uso:
	// 
	// Para cifrar (automático):
	// string textoCifrado = crypt("Teste 123");
	// Processo: "Teste 123" → reverse → "321 etseT" → substituição → cifrado → "#CM#cifrado"
	//
	// Para decifrar (automático):
	// string textoOriginal = crypt("#CM#cifrado");
	// Processo: "#CM#cifrado" → remove marcador → substituição reversa → reverse → texto original
	//
	// Gabarito de substituição:
	// Original: "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,!?;. !@#$%¨&*()-=][:<>"
	// Cifrado:  "WLygjSDROVAvbExsh;#itkMYcXwNU!Tzploa,CZIndK?JPBeQGfuFHmr$)(<¨%:@*!^q&-].=][>"
	private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        // RADIO BUTTON DO MÉTODO PADRÃO OU DIFERENTE
        private void rbmetodo2_CheckedChanged(object sender, EventArgs e)
        {
            if (rbruimetodooutro.Checked) tbruimetodooutro.Enabled = true; // Habilita caixa de texto
        }
    }
}
