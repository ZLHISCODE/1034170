-----------------------------------------------------------------
--ÎªÅäºÏ²úÆ·°æ±¾ºÅÓÉ9.31ÉýÎª9.32(VZLHIS10.21.0)
-----------------------------------------------------------------
--ÎÊÌâ:12796
DELETE  zlparameters a WHERE NOT EXISTS (SELECT 1 FROM zlsystems  WHERE a.ÏµÍ³=±àºÅ) AND a.ÏµÍ³ IS NOT NULL
/

ALTER TABLE zlParameters ADD CONSTRAINT zlParameters_FK_ÏµÍ³ FOREIGN KEY (ÏµÍ³) REFERENCES zlSystems(±àºÅ) ON DELETE CASCADE
/

--12617
Insert Into zlParameters(ID,ÏµÍ³,Ä£¿é,Ë½ÓÐ,²ÎÊýºÅ,²ÎÊýÃû,²ÎÊýÖµ,È±Ê¡Öµ,²ÎÊýËµÃ÷) Values(zlParameters_ID.Nextval,-NULL,-NULL,1,20,'½çÃæÇøÓòÒþ²Ø',NULL,'1','ÉèÖÃÊÇ·ñÔÊÐí½çÃæÇøÓòÌá¹©×Ô¶¯Òþ²Ø¹¦ÄÜ')
/

--ËµÃ÷£º¸ù¾Ý²ÎÊýv_OutNum·µ»Ø¼òÂëµÄÎ»Êý£¨1-40£©£¬Ä¬ÈÏÎª10
Create Or Replace Function zlSpellCode(v_Instr In Varchar2,v_OutNum In Integer:=10)
  Return Varchar2 Is
  v_Spell   Varchar2(40);
  v_Input   Varchar2(1000);
  v_Bitchar Varchar2(2);
  v_Bitnum  Integer;
  v_Chrnum  Integer;
  v_OutMaxNum Integer;
  v_Stdstr  Varchar2(50) := '°Å²Á´î¶ê·¢¸Á¹þ»÷-¿¦À¬ÂèÄÃÅ¶Å¾ÆÚÈ»ÈöËúÍÚ-ÍÚÎôÑ¹ÔÑ';
  v_Chara   Varchar2(2000) := 'ß¹ï¹åHàÄïÍæXÞßàÈÜtþH×cö°ì\íÁàÉæÈêÓè¨ÙŒø×rèP÷oìaèñâÖÚÏÕYì”ÖOéœõcùgíù“ëˆÛûï§ë@Þîä@áíØtØåB÷öálÛêÝEëJà»âÚéáåÛÖ’÷¡÷éö—úqüÞÖæÁéOá®æñÖ“öËðÆñúòü';
  v_Charb   Varchar2(2000) := 'á±ôÎášØ^÷„ôƒÜØá—ÝÃÝRïT÷Éü–îÙâZÚ•öÑõEå±êþÞãßÂì‹Ù”ívîCÞnÚæÛàîÓô²âkô‘é›ã[ì‡Þkäºß™íDÝòÖræ^ÙèæßìÒöµé–ý_Ýáï’ï–øRødÙ…ìdõÀè˜ãEìsõUètÚéùlãmØÚýã£àfÝKíÕÝíÕRÝ…ä^÷¹öÍêÚßGÙSï¼åQÛÎÛÐÝ™éaàÔìžßJÛMê´éGçaØP÷”æqùSösÝ©Ø°ßÁåþïõÙÂô°Ø„×ß›î¯ßÙÜêáùîéæ¾âØÝÉé[é]åöã¹ÙCÚPääõÏãGésïàŠæÔÞµõIå¨Û‹í@÷ÂèµàˆôÅç@íSí{ÜKÜLôxÚFèEúzú‡ü„í¾ìÔß„æQöýß…öböcØÒÙHíÜøuÛÍâíãêÜÐáŠÞÕçÂérÞgîYÞlÞmÞpÞq×ƒìáè¼ì©÷ÔïRæôûïÚì­ì®ï[Ö€Ù™çSïðïjïkïlïnèsæ»Õ•ål÷§÷B÷Mü‚ý–õ¿ß“ÙÏçÍéÄØhÙeÙfïÙáÙìEè\î éëë÷÷Æ÷ÞôWÙûä‰ÚûêvÙ÷âãuìïžðVÞðÕ@õmìhâÄà£ã\ðGÜ@÷QØÃàRÙñîàâ“ãK÷ˆõÛäcéDõNØmænùPíçè}õËô¤ë¢éÞ×LåÍîßêÎâ˜ÕcðJÞKõ³ß²øGùLûQîÐê³âbÛYà^ð¾ñ£ñ­ñÑñØñÛó÷óëóÙóÖòùñÙñÔñ¹ñ¦ð±ðÇ';
  v_Charc   Varchar2(2000) := 'àêíåßnØ”ÛPï{æî÷õüoôÓè²ÖØ÷û]úIè†Ù‰àÐäîô½ç[Ü³à“üâüá¯ä¹àáè¾âÇã˜ïÊåšæ\âªìxé¶éßÛ‚ïïèdãâæ±îÎâOÙ­Þ{êèäiæ¿åîìøÕSäaâÜäýàšïâàžõðéK×‹èÚÆÝÛÕ~éˆÙæápÖçPêU×€âãåñí]îØöæ½ÝÅãÑè å_é‹öðöKüÜÉéLéMáäæÏä–÷•çL÷lêÆã®ë©äâêÛËÕkíoâ÷ìÌânêËà}ü…ÞCü{ÖšûžÜ‡íºåøÛåÞŠîJÞÓè¡àÁÕ€ÙoÖnÞå·êÚÈÜ•â\ëÖRû‰úmÚ’í×Û{Ù•ö³Úfé´ýYýZÚß×êpèßîõÚWìlîªÚXîdçdçpèKØ©èÇàJÛôîñëóÕ\õ¨ä…õ“ßêí÷àÍæÊÕvø|ùA÷ÎýcüJü[ÜÝÚdÙPßWÚmßgÜ¯õØßtÖsôùãrýXáÜß³âÁÞŒë·à´ï†ÙÑÛLãMë†ßoã‰Ú†ùúuâçÜûô©ã¿ô¾ÛŒê™ï¥ã|Ù±àüã°áOëlöÅÜPá~×‡×‰áhô{ßcéËØŒýiÛ»ØaÚnãIäzërõéúRÜXèÆèúýsýƒØ¡âðç©ØXàsÛUézÕ‘ãÀ÷íÞõà¨àÜõßçÝë°ô­å×ÝŽâ¶îËâAÙiúEêJâëý—Úïé¢é³åNæmîqÝöjùœêÝ»ácåTõžù‡ÙƒÛwõÖåÁÞußOê¡áQÚ}ÝzöºýpèqýwßÚÚeìôÜëÞeâ‘ôÙÞiï“ð@øyÞoú\ú]ÚÙnÜÊèÈæõè®çWäÈçýÕpÙzÙ{Öé¨ëíê£Ýû€û‚û›áÞéãâ§õ¡ÝýÕKÚuÛqõ¾üyõíÜAî•Ùàß¥ïéÜfè‰ìàéÁçJè­Ú~ßýã²ÝÍë¥îxß—ÛZââßuõãáiáÏïóõºûzý€ëâØÈßHï±äSåeå¤ðûö¿óøó×ò¿ò²ñéñåñÝñÒñÎñÃñ¬ó¸ó©ó¤òíòÜòÉñ¡ðîð·æöðËðÈëú';
  v_Chard   Varchar2(2000) := 'ßÕÞÇàªæpÞ…Þ‡æ§âòí³ßQß_ÛQ÷°÷²æ]ÜJèNí^ý‘ý“ß¾Þaá·ß°çªåÊçéÜÜ¤Ü–ÙJÝDõ\øl÷ìÛìOünì^íñÜláGééàîFÙÙü^à¢ÝÌÕQå£ø}ÙœìKÚÔ×[üh×•ÛÊå´í¸ÝÐßTë‹ÚêWØÖß¶âáë®á’÷ô€ê‰þOëIëZìâÜ„ôîï½åuØOàâô£ê­à‡ëQáØíãïëç‹ôÆêÚhàÖïáå~íLçCÙáÝ¶êëì{îEô†ØpûMØµÚ®Û¡êsÛæèÜíÆÝB÷¾öWæ·ÞžÚÐâKé¦íûßfãdíÚßrÖBÛyàÇÛ†áÛîŒîý‚õÚücÚçÛãçèîäâšëŠô¡õõøJõMöôü—õ ùmážîöâyäHëÕ{ä”èSÛìà©Ü¦ÞéÚgëºéPÕ™õÞöøölØêçàôúá”ìwí”ü‡ç–ï}à¤ëëíÖåVîrîûïMäAßËá´ë±õ[üŠöCù…úHÕ‰Ûíá¼ëËÞ“ëØíÏëšÝúêhî×þKâ^àKôYáHékôZäWðLêLô^ô`ôaê^à½á`êAäÂèüë¹ÕiåL÷ò×xØKÚGí~÷Çèoíbíüt×˜Ù€Ü¶ì|åƒæHé²ìÑå‘ÜYîXø‹çŽí¡êŒê íÔí­ïæ×Bç…×míâÜHíïõ»ÜOãçìÀí»ÞšâgîDßqÛvßÍîìâ‡õâõyèIßáç¶ÚrÜoÜ€ôDãõêwêyÛFÛGï˜ùzð¬óýóûóìò½ñõñôñóñ×ñÖñÉñ¼ñ²ñ°ð÷óÎóÆó¼ðãðÛð´';
  v_Chare   Varchar2(2000) := 'åíÞˆÝ­âeï°ÕMä~îPô‰î~ùZù[×Fæ¹ùEêißÀÜÃêqéîÛÑÚÌÜ—ãÕãµÝàØ`ÝQß]ëñïÉß{îOðIØ¬Ö@é‘åŠöùî€ötù˜×†èyý|÷{ÝìÞôíEêzÝ[öÜëXõbøÞWåÇçíîïãsðDßƒÚÙ¦Ù@ÙEð¹ò¦ðÊ';
  v_Charf   Varchar2(2000) := 'áeÛÒéyíÀåzá¦ÞNïcïx÷YâCÞ¬ìÜõìÞÀçxú‹Þxî²ÜèóØœÝGïˆï‰ØÎÚúèÊîÕÚ“â[åpøhöÐô™áÝô³úJåúïwç³ìéìqö­öîöEïyäÇëèã­ì³é¼ôäÕuáôÙMïÐü”çšì]çãÜmâpëƒèûëVôšøXØk÷÷ü‹ØrÞMèMüRüvÙÇö÷å¯÷aããí¿ïLÝ×à•ähØSæ‘çQÛºìbïpüKßôÖSÙºÚRøLøPøiÙˆë€ø]ß‘ß»õÃáKôïïûõÆâaà~ØføWûŸüAüFÙìæÚÜ½ÜÀâöç¦ç¨ÜÞìðî·ÜòÛ®í‚øIíÉÝ³ÙëèõåõÝÊþEãRãVïOøDíhá¥øqÖDÛ~Ý—õHõvíêùfù›ß¼ÞÔàMáœäæÝoôfíëÚâæâØ“ê‚öÖêçÙxÝ•õVÙŽå‡å˜öûövð¥óõò¶òãòðó¾òóòÝðò';
  v_Charg   Varchar2(2000) := 'ê¸Ù¤îÅæÙáåmæØÞÎôpà@ÚëÛòêàëBØdÙWÙ^æYØ¤â}ê®Þ|ÛáãïÜÕôûÞÏøNôvýžä÷ÚséÏß¦÷ ÷hêºí·ç¤äÆÚMêlî¸âGä“æsí°éÀØºízúküŽúê½çÉéÂÞ»æ€Ú¾Û¬ï¯Õaä†ÛÙæüéxømøwÖgøæŠØªàÃÜªë¡ëõéwïÓì‘ík÷ÀÖYÝ‘õsækíuÞPíRöÛÁô´íÑãtßçØ¨ôÞÝ¢âÙûfÙsùˆàQßìç®öáõ†ëÅö¡ÜpýŠýÞÃçîÝ\ì–Ø•ÚCØþçÃâhã^÷¸íxá¸èÛØxÚ¸æÅì°åÜëgêíÙéïÝÔõýÝLÝMôþì±âõYøÝž÷½úXãéÚ¬êôî¹ßEîÜâ’ü‰ØÅëûî­ù]áÄèôêöïÀíådöñöAî™ëÒïNÚoäTïWøŽßÉØÔÚ´ÙÄévêK÷¤êPöŠ÷bÝ„å]ÜIøAÞèäÊØžßkîÂëqæšè…ûX÷}ßÛèæë×Ý_ã üUáîæ£ßžàFé|öÙõqý”ôhôkþIå³âÑØÐê{Ü‰êÐØÛêÁÙF÷¬íW÷Z÷iØ­ÙòçµíÞÝöçõPõ…ÖßÃÛöáÆâuåàþÞâë½Ùåâ£é¤Ý{ðRèJß^óþóôóàóÑòåòäòÁò¼ò´ñøñæñËðáðÙðÀðóð»ð³ð§';
  v_Charh   Varchar2(2000) := 'îþãxàËëÜáVõ°ï™í™ØEõA÷ýÚõêÏìÊäwínØJô_ê\ÝÕâFé\ÞþäIädîhîuÖ›ënå«ú[ôŒÞ†ç¬Ø˜î@ãìÝïàãÞ¶àÆå©×qê»å°î—ö‚Ú­àÀÛÀàAêÂîÁý†Ø€ãFãØ÷…éuûiûîMôçôŸêHíHý[ùŸèYý˜ëaÙRÛÖúQýLìeìfûSìgü\ì•èìçñûaø’ùCÞ¿èUÙêÜŸØFÞ°Ý“åÞZãÈãüÝ¦ØAâvébØDãpìô„äfÞ®ë”ÙäëŸø™üZÚ§é{äUé•é—ô\ýJãô×÷¿æAö\àCááåËàjÜ©Ø_ö×÷õ`÷cìÃéõßüã±Ü ëŒäïëiÖ—àñõúâ©ìÎéÎô–õ­îgì²æLôEö{ù–úCúKä°çúåtöUÙüá²âïìæìïìèà‚øUå×o÷Ÿí_í’÷sûIÕjåkæèîüänÖœçfú†èëÕ–Õ üXõ×øbùJà âµØŽ×’Û¨ä¡ÝÈëfØ}ïÌêaå¾çÙß€ØoæDéIûqÞSêXèG÷ßÝkÛ¼ä½åÕäñöéß§õŒöZödëÁÚòüSáåäÒåØäêéBè«ÖWå–öüÚ‡í‹çuöm÷UúŠÖeæwÚ¶ßÔêÍçõØYëDÝx÷âãÄö™ä§ÜîÞ’ßDõt×eßÜä«Üöí£èíåçà¹çÀê_ÙVÕdÞ¥ÖMî_×M×fçiêTçžìuí}×wîœãÔé’âÆðQÞFý@Ú»äãÕŸïÁØååxß«ß˜îØâ€â·éXØ›þAÖfëoïìàëÞ½èZì[ð©óóóòòºò³ò«ò¥ò¢ñþñüñëñ¥óËó¶ó³ó¨òÂòÀðúð×ðÉð­';
  v_Charj   Varchar2(2000) := 'Ø¢ß´ØÀçáÜ¸í¶ßÒØÞßóåìï|ïúê÷êåõÒøKã‚çÜÙ}Üuì´ÛÔåZëYî¿ÙŠàœëu×Ií‡ù×^çˆÜQíZúaýVèWèiýWûAá§Ø½Ù¥àBþLØCê«éêé®ÝðÚlãšÞªÛeì“ûnÝ‹Ûˆå‰ÞUçgìPúWúnÜeë|ë}Þá÷‚êªáÕ÷äô‡åæÜÁßâä©êéÙÊÛEëHôßÕHõÕö«öÝÕ‚öê÷ÙõJÛ”öaùHýTæ÷õŸ÷DìVö›÷C÷qåÈä¤çìôÂÝçõÊãeïØØjØ†æ‰û“áµÛ£àPí¢ê©îòÛOïäeî]îaø”ùGëÎÙZâ›ê§ÝÑØ]äÕêùégìyÞöçÌÝóØböäøZä’íKû…÷µ÷œùpöx×töúYí[öžè~ídàîèÅõÂíúïµÚÙê¯ôååÀå¿ÖˆôCörûxç‰ç™û{×vû|êðÚÉâVé¥ë¦ëìÚ™éfÙ`ÙÔÕÙvÚ{Û`õÝÖGæIðTæGçZÞYèaèbè{èƒÜüôøçÖ÷šíäí\÷FÖvîŽä®ç­êñánôÝáuÖ˜Ü´æ¯ÜúõÓÙÕöÞõoøŸÞBç€ú„úŒÙ®ÞØäÐë¸Ù]Û]ãqïœáèùa×K÷Rá½Ý^àÝÚŠÞIõ´×_á†ëAà®àµìŒù™ÚàæÝÚ¦ÚµÞ×Þ—èîæ¼ã]ô‚íÙöÚôÉÕmÛdîRæOõ^ï÷ºÕ]ôîÄáŽûvüTÚáÝÀâÛâËéÈèªå\Ö”æ¡Ý£êáßMçÆêîàäÙÚBý„ãþìºÝ¼ëæùXöLù~ù‚ûü û—ÚåØÙëÂÙÓã½îiåòåÉëÖÞŸæºö¦â°ÕeÛVîKìnìoçRØçìçåÄÞ›ïGîyôñãÎà±øF÷ÝôbéNíƒèÑèêÙÖöJûýnúÜÚêÞäé§è¢ôòÚ ï¸öÂÕ‡Ûgä|õLø~÷¶ù‰à`Ý]ÛRÚzÜvé…éÙùVÛžùqúGüŸýAÜìé·é°ö´þFõáýeÚªÜÄßšîÒÙÆêøÛBâ ì«ØeåðõXåáäïZöÄØ‹ÜMõ¶èLä¸ägämïÔæŒùNçîÃïÃäŸèðáúöÁÛ²ëhï…ðCàÙæÞçåáÈèöõûÚbÚ‘ßIØÊÚkâfØãÚÜâ±Þ§ø_ø`àåéÓéQïã×HõêÜBùŠÛÇç~çìßú€ý™ØÜjè‘ÜŠâxãzã—÷÷åå‹õzûŠûŽê}ÞÜðKùQùRùUð¢ðÏðÜðýóÞóÅòÔòÌò»ò±ò¡ñäñÕñÐñÊñÆñÀñ¤ðèðÕóÕóÈóÇñððÔð¯ð¨';
  v_Chark   Varchar2(2000) := 'ßÇØûëÌãlï´é_ç˜ØÜÛîâýê]îøÝÜÝaïÇå|æzêGïaâéæbíèê¬ýÙ©Ý¨Ý|ÝîƒÞRãÛî«êRÜ{ç_÷KØøß’ãÊîÖâ‚é`åêèàîíêûäD÷Šõwõ‘çæéðÚîÝÝVïýâŽî§îWáfîw÷Áá³ã¡ë´æìç¼à¾äÛï¾Õnä˜ØcØ~åoï¬ÕUäLå”çHÙÅáÇÜwÜxåIùyìÜÒíîßµâ@Þ¢údØÚß Ü¥Úœ÷¼õpç«à·ÕFÙ¨ã’ØáÛ¦ßàáöëÚ÷Žà”÷d÷ÅèwÚ²ßßÑÕEÝHÚ¿Ü’ÜœÕNù\ÞÅÚ÷ÛÛæþêÜÙLÝAãkäqà—üYèkã¦êNîåÓàkí–Ø¸à­ÞñêÒî¥î`åžæKÙçÜiõÍíŸÛ“ØÑà°ã´ÝÞÖdçqè^çûï¿÷Õûdõ«åKöïöHù{úAã§ãÍé€éèéîSéŸíAípìHíTôUðâòÒóñóíóØòòòñò¤ñÌñ½ñù';
  v_Charl   Varchar2(2000) := 'ååê¹íÇØÝÞhôFéJö_èníBáÁáâäµà[ßFïªånöDù„üHêãíùÙläþÙ‡îmîsù`ô¥á°ìµïçê@×E×ŽÜ_è|è”íeé­äíî½áYà¥ýœàOÝ¹ïüï¶àHÜqäZæƒãÏÕLéÝõßëáÀï©õ²ç„î‘èáîîã™õuÞLÜ~Øìêbß·ãîí‰÷¦ö˜ðEæÐçÐéÛÙúèDÞ[èhìY÷mýFÚ³ÕC×|èˆûPõªãîLî[åGïKîàÏÜ¨Ûkã¶æêà¬çÊÝñæËØ‚ä‚öâî¾äœÖ‚árÞ¼ß†áëxõ”ç\öPùv÷óèg÷~ûZþGÙµæ²åÎï®ØNýŸä‡å¢õŽõ·÷¯ßŠ÷kß¿ÛÞÜÂìåèÀÙ³èÝÚ\éöÛªáûíÂÝ°à¦ôÏîºõÈö¨äàãWøEë_øtë`ûáBúbû•ÜVÞ]×Þ^ìZ÷uìcÞÆßBöãå¥ì¡ÛšÖ‹æ`×`ôHç ö–çöÝüà˜æ®éçé¬äòåbå€æœönýé£ÞcÜ®õÔÝˆ÷ËôuÝgÕÝvåyÜGàÚå¼â²çÔß|ØIÙ’ÛŽç‚ïmúîÉá‘à€Þ¤éRÞÍßÖÙýä£Þ˜ÛøÞæôóïVõhø•õñôQ÷à÷vßøôÔàëOý á×åàê¥î¬û‹ÞOçl÷ë÷[âÞãÁéÝïCÙUÝþì¢éŠÜCõïÜ\ÜkÞ`àòãöÜßèÚê²û_èùç±ôáÚšÝCâéqÝsë‘ä™ë™õCöìøoûwëëžýhÛ¹öNýgáì`û™ý’êtîIßÊìÖä¯ì¼åÞæòïvïÖûméHöÌæyûˆçBïdçsïiö†úVç¸ï³äÛ‰ìCëwïfôjúwãñÜ×èÐççëÊíÃýˆçXìNýýŽØLÜ[èxì_ûTÛâë]ÚLÙÍà¶ÝäßsÖŒÜ}÷ÃíVáÐïÎçUààß£ÛäãòèÓëÍéñôµâ„öÔô—Þ_èzïB÷|ûRüuûuô”éÖïåæ”çœèuéûê‘äËåÖÙTÝ`äõÚ€Ûjê¤áXä›åhåjè´øšÛÞAçGöIùcùnçeú˜ëªãÌéµé‚úyàLïùëöäXèrèïÙõöÇùFèŽû[á›ï²äsäxàðêÛiÝ†ä—öMÕ“ÞÛîbâ¤ëáé¡ïÝæ ß‰úŸèŒÙÀÜsÙùãøÜýçóÞûäðöÃõiðµðÓóüóöóÒó»ó¹ò÷òëòÛòÈòÃñöñ®ñªñ§ðüðøðìðßðÝðØðÒð½ñïñìñçñÜñÚñÏñÍð¿';
  v_Charm   Varchar2(2000) := 'æÖáïßjæ‹úiö‡è¿éUßéôKö²Ý¤ÙIú”Û½ûœÙuß~ì@ìAî”÷´÷©ôMôNö æžà„Ü¬á£çÏì×ïÜÖ™çNÚøíËâIèšä€äÝØˆêóì¸ÜšáF÷Öå^ùšá¹ã÷ÜâêÄãTë£ÙóÙQà|è£î¦àŽí®Ý®àdáÒäØâ­é¹ïÑäYæ[úBüqä¼ÜzæVüeÚ›ômíi÷ÈÞÑîÍéTéYå{ìËí¯ë‰ÝùÞ«à‘à–ëüíæõ’ô¿ûsìXîŸûLÛÂô»åiãÂöQü€ìDìWÛ_ßäìòâ¨Öi÷ãû†÷çû”ÞÂéSá‚áƒûJáˆØÂåôôÍëßãÚ¢ôéãèåµÚ×à×ü†Ökå²ãæö¼ííäÏëïõ|ìrû ü@üMüIß÷ù‘÷]èÂíðíµç¿åãØ¿ßãøpèf÷xáºçäÜåçëçÅâŒÙ‚ä øsæFãÉãýéhíªüwé}÷ªöšÜøÚ¤àpäéêÔã‘øQî¨õ¤çÑÖ‡ÚÓæÆâÉüN÷áôžÖƒÖ„×OüOéâÜÔï÷Ø{Ýëõöã€ì…ïÒôŽüaõøæŸßèÙ°íøãwÖ\öÊøœüEë¤ãaÛ[ØïãåÛéÜÙîâë‚ãfëŽíJðÅñÇðÌó·ó±ó¡òþòýòúòìòÖòµóúóºò©ñòñ¢';
  v_Charn   Varchar2(2000) := 'ÕyïÕæ“ë~ëÇÞàØvÜ˜Øyâcì„ô›ÜµÞ•áèÍÝÁØ¾åràïà«ßaéªÖQëyôöëîàìôTâÎêÙß­ýQØ«ßÎíÐîóâ®×DçtÛñè§émô[Ú«ðHõƒõàÅâ…äGÛèâõà\îêâ¥ÛCâ‰ØƒÝröòöFûŒýuÙ£ì»ãbëWèXÞ‹êÇíþöÓõRöóùDöTéýÝ‚ÛœÜTØ¥Ûþá|á„ÜàôÁøBæÕëåí±Úíô«êŸà¿ãcÛWÛfÛhåRõææ‡êEÞÁým×‘Übè‡ïDèßÌè_ôVûHØúå¸æ¤âîáðâoìÙ¯ßæÞrÞsáxýP×aæeç×kæÛæååóæÀîÏâSí¤ô¬üQàGÙÐßößSÞùï»ÖZÛåŸð¤ò¨ñ÷ñññÄòïòÍ';
  v_Charo   Varchar2(2000) := 'àÞíMÚ©ê±økÖŽæ–útý{âæñî';
  v_Charp   Varchar2(2000) := 'ÝâèËÙ½Ý‡ßßÝåæWãÝÛAõçÛ˜æoíQãúîGäƒùbè‹ë„äèìQåÌ÷›ý‰ý‹ö„ëãâÒáóÞËÝNìŽûƒüBõ¬êkêŠïÂÙräžàúì·àÎö¬Þ\äÔâñÝJéoàØÜ¡Ý~åAíŠíŽùiôJèmÛsêCØ§ç¢ÚüîëØwâWâtâ”ãYãàèäšåCõBêVêoÜÅèÁÛ¯ÚðÛýØu÷‰î¼ëRô“õQõùùdÜ±âÏØòÛÜã›Õ|øaß¨äÄæÇî¢ê¶úûGêúôæú@æéëÝÙXÕ—õäÚÒÙGÕ›ôØâçÎïgïhôwêQéèî©áoî’àÑæÎë­Ø¯ÜÖçvæ°ØšæÉîlïAé¯êòæ³Ù·îZàZèÒÝZöÒÝƒõGîÇá•áNáwçkÛ¶ÖcØÏîÞãOçêîHïHÞåÙöê·ë¶ê†àÛäõ‹ÙéáTè±å§ïäÙŸçhäßë«ÖEïè×Võëç’ð«ñâñáó¦ó²óáóÍó´óªòçò·ò­ñÈñ±ðå';
  v_Charq   Varchar2(2000) := 'Þ€èçàVÝÂàÒéÊÕƒÛpÖ[ë’õèôtçKù†ØÁÛßáªÜÎêÈä¿Ý½Ú–Ü™âHæëçùç÷ì÷þDèŸí ônôoôëýRÞ­ÛaåW÷’÷¢õšùuù}÷èôGôyö’û˜ßŒá¨Ü»è½ØMç²ôìÖHêMãàÜùÝÝíÓí¬ÝÖÚžáMì—÷ÄÚäÜ·ÙÝá©ã¥Ø@âTâ`ëeí©ãUå¹ûeåºÕßwå½ÖtîvçcùkèBôRôSíaÝ¡îÔÞçÜâjã@ãQäEåXæZübö‘ÛÉëÉã»ç××lècÜÍÜçÙ»èýÝ€ãÞê¨õÄïºïÏäÛ„ïêÛ–æjçIçjæÍéÉÖmôÇìÁíÍàbàzõÎàƒà…ØäÛ^îNçØå æ@Ú‰ÜEÜFèAÜñÚÛã¾÷³éÔ×SÚˆçyíXî˜á ã¸Ú½ê~ÕVímíIÜNæªêüã«ïÆôŠÛoå›ö@çƒôÀÕWîzõÜËâsëdàºäÚì€àßøVéÕÚ_Úcï·äußÄÞììiàWàõÝXÝpöëõ›è[éÑ÷ôÜÜí•Õˆö¥ìmíàõ¼öÆÚöÜäÚ^é±ûjÚ‚öúíFíGöpöqù”÷Gý•áìÙ´ÞåÏá–êäâUÛÏåÙôÃÙgäMábõF÷üõ‰ùj÷AôÜá«Ú°êrìîÕoüLõ@Ú…üDÜ|ôð÷ñö÷OÛ¾ëÔÝ@Þ¡íáøzè³üšÞ¾ë¬áéÜdèŠûYýxÞ‘àTãÖêïé‰üCé˜üzãªçzÚ¹ÜõéúîýÛIÝbãŒÛmêB÷™÷ÜöeýjïEáëî°ç¹íjí¨ã×ãÚÚ|é êIùoåÒð¶òøòéòàòßòÞòÓòÐòËòÇóéóæóäóÜóÌóÀó½òûò°ò¯ñýñûñßñ·ñ³';
  v_Charr   Varchar2(2000) := '÷×ÜÛìüÜ`ôX×j×ŒÜéèãæ¬ëNßvØéâmôã…øžÜóïþÜrØð×šéíâ¿í¥ÜÝØìzì~ígïƒÕJïšÞwê—âJâ~ëÀáõáÉéÅéFægÝPôÛõåÝŠåˆ÷·ökù’íqßï¨ãœønàéå¦Þ¸ø›á}îž÷pàrÞzä²äáçÈÝêøMëÃÜ›Ý‰Þ¨ÜÇèÄî£äJä„écétÙ¼àeö}ö”úUð¦óèòîòÅò¸ò¬ñÅñà';
  v_Chars   Varchar2(2000) := 'ØíìƒØ¦è•ìªëÛâlëMïSàçî|öwÙë§ôLôÖâÌçDédÞúíßærî‹çÒëýöþïbö…÷fÜ£ØÄï¤ëäCÞQçmÖ ïoôOé~ï¡ôÄô‹öèéŒæ|õõßþì¦é„ö®áêßÜÏæ©îÌÛïô®Ü‘áŸé^õÇäúëþõŠê„éWÚ¨ØßÚ]ãˆæóÛ·æÓÖb×iÙ ç—÷­÷W÷XéäõüìØÖ…ôlÛðÙpèlç´ô¹ÝiïYõ}ÜæÛ¿äûâ¦î´ÝfÙdÙhÙÜØÇäÜís÷êÚ·÷“õ˜ù_öYöŸãhö•ß•ßÓïòÚÅ×ŸäÉÕ”îTô×}ëÏé©ävêjê…ê’þJãHå•ü›ù|äÅ×W÷jíòêÉÙKáÓÙ‹ßŸû\âPÝéãAõ§øOø[öõåœöXö‰úPá‡â»ÞyïzìÂÛõÝªßYãJãvöåõZüœ×Rüöˆõ¹âìêêÛéøîæá‹ß±ÚÖÙBÝYâ‹âžï—ßmäKÕœÕžß}ðSÖuáŒö|ýaÞÐô¼á÷ç·æì¯ç£êxæ­Ù¿ÝÄÜ“àgÞóë¨ÛSÛ\Ý”õ_ùeïøÛÓÚHü“÷núž÷tãðëòã_äøØQåfçTùŽùà§ÕXãÅéVäÌþBëpæ×ú{ûtûUç`Õlãßéjí˜ôBÕfÕhåùîåàÊÞ÷ÝôéÃælèpÛÌæùßÐçÁãjïtØËäFïÈäùälæJï\çrúƒýDãáÙîæ¦ìëãôæáÙ¹ï~âLØ|â–ï•âìÚ¡áÂäÁÝ¿áÔã¤ížÕbæànà²äÑâÈì¬ïËágæ}ï`ÛÅàÕî¤Þ´öÕõ‡Ùíä³ÚÕà¼ãºßiûhÝøö¢Úxßpä_ðMÖqÛ‘÷Tú‰â¡Ý´íõî¡å¡ìšëmßUëSÚÇÕrÙwìÝåäçw×\ç›áøÝ¥â¸ïŠöÀé¾æ{úZæ¶êýèøíüàÂôÈÚtõ€ßïæaæiæææ•ßCð£ð¸ðÞððóßóÓóÏóÂóµó°ó§òôòÙóùóâò×òÏòªñêñµðþ';
  v_Chart   Varchar2(2000) := 'õÁîèäâãBÜDõ]÷£öãËßeåÝê`é½ÕwåJìŸêFíOêY×nÜcææÛ¢ìÆõÌöØïUõTÞ·ëÄîÑâØê¼Û°ïÄÕ„á]åUît×TØáv×Zú‚ìþîããgáaêæÙyï¦ôÊïÛÛçMç|íUü‘â¼àoäçëGè©éÌÛ}ÚZõ±æhêOúSàûÙÎéEæ†è’èºï‘ÖzíNíw÷ÒÞä¬ßûìŠá[ì’ä•åcØ»ìýß¯Ø–ï«í«äˆü’ëøß‚Ö`öŒìLäRúeúfç°ç¾ßXÚ„õ®ÖpÛ‡å÷–ø˜î}õ{ùYö[ù•ù—ÜnÜƒÙÃã©åÑßPç‘ÚŒáLìjüVìpî±ãÙø‰êDúcúlãÃéåï›Ùqå`ìtÞÝÙ¬ìöö¶äpì›÷ØöæõæxýföœÕAôÐï¢Úqî\ÝÆÙNÛ@âŸãŽø‡ç“ç”èFï”÷Ñî®ì˜ß‹ÜðæÃÝãéƒöªÖFüžèèîúïFÕPäbîcàÌÙÚÙ¡ÜííÅÚUãPÙ×ãnã~ï ÷‹äüõjâúæBÙï÷»î^üWäŒùWýCÝ±âŠÛTõ©å„ùIùúhú“îÊâQÞƒÜ¢ÝËùrØ‡ÞÒæ˜úoú™î¶åèëPîjîkînôsÛÛƒìÕêÕü`â½ØZëàÜ”ï‚÷ƒôë˜ÙÛØ±×™ï€ô…Ù¢êuÛçãûÞèÞíÈõÉõ¢Û|éÒõDørü˜ö¾üƒâÕùKözèØÚ—ðÃñ»óêóÔó«ó¥òèòÑñíñÓ';
  v_Charw   Varchar2(2000) := 'æ´ü|Øôßœëðící€áËî“ØàæýÜ¹Ø™îBßÝ¸çºëäÝÒçþîµÝnÛläjå†ä[åsÙ–æ~Ú@Øèã¯éþÕsÝy÷ÍÞ‚ÙËåÔêžÚñÝÚìÐÞ±÷˜ögöhàíÛ×àøãíãÇífä¶á¡áÍß`àŒáWå…éõdìSìTì¿çâä¢æ¸ÚÃÚóâ«ôºè¸öÛÕ†ÛcílîQå—õnítï]í|ê¦â¬Ö^åMõKÞEçAìG÷×~ÜZ×ˆÜ^ÞdØnÝ˜Ýœ÷—æ’ö€ö“ãÓâ†ö©ô•øYøjéé”éšü•êZê[ØØãëî‚è·ûlæfúOÝîÞ³ýNÙÁÝ«à¸Ûbë¿á¢ä×íÒö»ý}ÛØÚùÚâEàwÕGÕ_øŒæuöƒàNßíä´ûcùMõˆ÷ùú~ØõåüâÐâèâäåÃêõåqù^ÜRØ£þ@Úãè»ÜÌßAìÉæÄëFì}æðå»Õ`öÈëœýHìFýIúFðÄòêòÚðôðíðÍ';
  v_Charx   Varchar2(2000) := 'ÙâÚÀÛ­ßñÞÉä»ì¤äÀÝ¾ÚTâRôÑôâô¸àqÙÒÕOØgðFæÒðOéØì¨ìäôËåaØGØHØlØ‰ëvõ–ùT×@õµç^ë^êØá@÷ûú è„àEêêÚvÚôÖæˆìI÷žïe÷@çôáãÝßâ|åïÝûãŠìûÖLÖlÛ’÷^Ühâ¾Û§àSâMãÒôªÚiìùÚVëKü_ô]êSìUßÈØBéiïPöyáòèÔêƒíÌê˜åÚè¦ÚYô Ý å’æ_÷ïúTépÕ’ç]ë¯ììôÌÝ²èõÑõ£ã”åßí„åvåwí†õrÛŸ×]úNÜ]÷€æµéeã•ÕtÙtÖPÝá_ûyÚDèvú‘ú’úšÙþáýê“Ú`õÐëUÞºå‚ìÞî‡í`ï@á­ÜÈêˆÕ^ä}åDØRüGçoö±ýEÜ¼à_àlàmç½ÝÙàxæøû‘÷`è‚âÔÛKâÃ÷Ïã}ðAößõaõœí‘÷zí—ç}÷PèÉßØèÕæçç¯åÐû^äìÛXäN÷Ìø{Öyø“újàUáÅÕqÖjÕ[þMÙÉÛÄß¢çÓÖCíPèH×ýšç¥ÙôäÍéÇé¿í…âÝâ³Þ¯åâÛÆÖxíCå¬ý^ýkýKÜaõóâàß”ê¿Ý·âdì§ä\Ü°öÎôgç†êcØ¶ÜŒîˆá…Ö_õSö]ÚêàDè—ê€íÊè™â]ãoã‹ätß©Üôã¬þNÜº×›×œßÝâÓõ÷âÊã–÷ÛæTõxø æ™ïqá¶äåäPçVçnýMíìí¹çï× íšíœôqÕšÖžôzè`ôPÚ¼èòà†ôÚõ¯äªÛÃäÓìãÙ[ã„÷rÞ£ÜŽÚÎÞïÝæêÑìÓÙØÕÖXæM×Xö~×zäöè¯ßxãùìÅîçäÖé¸ãCíÛïàìœïXæ›ÚKÞjíYí´û`õ½ëzú›ÞG÷¨÷LÚÊÚpÖoÛ÷â´Þ¹êÖõ¸á¾âþä­ä±Ü÷à‰öà÷S÷\áßÞ™ÙãßdÙbÞ¦îšèRðªóïóãóÚóÁó­ó¬ó£òáò¹ñãñ¶ðïðçðÂð¼';
  v_Chary   Varchar2(2000) := 'Ñ¹èâè›øfåEø†ùsçŒØóá¬çðíýý\åÂÛëæ«í¼ë²Þëý…âûëÙáÃÝÎäÎëçÛ³æÌáZéŽüiÚ¥ãÆåûÜ¾àIéZééÜîî†û’û}ûšÙðÙ²ÙÈØÉÛ±áDçüßVëCî»÷ÊÜyüdöoùžüfýdýŒüjükî›ýBô|÷úýzüsêÌêšìÍäÙøHõ¦ÚÝ÷ÐøeÖVØÍôeø‘àÙžÜ‚ázú`ÚIÚJ×…á€úŽá‰×—ØVØWãóãZë‡÷±å}ø„êgì¾è–ïráàìÈê–ÖUÝŒåø—ï^ç{öuìRûFÝIðBâóí¦çÛØ²ßºø^Ø³ëÈé÷çòÝUáæßbã“ïuïŸðPôíÖ{Ö|æc÷¥ï_î–öŽèÃáÊø€é™ýoúrì‰ýGê×ú_×Šè€ÞÞîôâXäyæUÚþí“êÊìÇÚËàvà’ØÌÖ]æEædùwìvûEÞvßÞâ¢àcì¥äôãžàæûpát÷ð×búsüpÛÝÚ±âùåÆâÂß×ÜèêÝÞ–íôôýÙOÕBÛDí›ïßzîUîVáÚî{õkÖ–çF×‚û@Þ~îÆÜÓô¯áÞ ârãiøCì½Ý}î‰ÞTýtß®Ø×ØîêdØýß½Ù«á»âøæäÞÈÞÄôàéóã¨ÞÚØ—ê‹ÛüØ[Ø\âNëcÚ˜ÝWâzçËìˆìÚÕxï×ûkûoü]ØæéìÞ²ôèØŠõlÙ“æ„ïîØsìJöGù€ùù‹×g×háyá{èOú^úgÜ²ú…úœ×”ý~êfä¦ë³êŽî÷ê”ê›à³Ü§ãŸìÖNë–éžë íÛóáþâwý‡Û´â¹Õzãyö¸ö¯ý]ýlúÛÈßÅâYâiï‹ëLì‚ï‡Úyë[×ÜáØ·áSõgÝºçøéAàÓÞüÙaè¬ë›øŠâßíŒævúD×súLè]úˆÜ…ú—ûKûWÜãÜþÝÓéºäÞÝöäëÙøÖhå­ÚAÛ«ïIîeçëôì™×Gà¡çßà{Ü­ã¼äVÛÕàaïÞëtçO÷«÷Ó÷Iúxà¯ïJî„öÙ¸÷‘Ûxõ—ákØüßÏû~à›ÞÌèÖÝ¯ÝµÞœà]ßKß[öÏéàâ™÷†ÝjôœÝ’õOßˆØÕÝ¬îðäBë»÷îÙ§àóå¶Þ”ØzáRÕT÷øæúÞ}ê|ßŽæ¥ì£ì¶ô§Ø®ô¨áüÚÄáCâÅÝÇâDêœö§ô~áÎÞíëéè¤êìÕ˜ëkðNôˆÛuÝ›å“Ö~õ‚öVöiú}ûCØñÙ¶àôàöâ×àhèžÕZäoö¹Ø…û‡ýrí²åýâÀêÅîÚþCÚÍßNãÐï„ìÏÝ÷â•îAØ¹ÝhãƒëTßyä`ø\ìÛÖIå[é“øƒø…øˆôráqùO×uÞXçŸìMå÷÷Nú–ûOÜ†ôcôdíóûgøSä‘øxùtüŒüÚOß–ãäë¼Ø’áJâƒö½Ü«ô’à÷éÚÖwÞ@üxæ…ß‡ù úMßRßhÛùæÂÞòè¥îŠë¾Ü‹îáÚ”â_ãXé†é‡éÐÙßå®ügÜSûNý›ûVÚSîfÙšç¡Ü¿êÀàyë…ëµä]êmáñéæâqëEìBýqýyÛ©ã¢àiã³ß\è¹ìÙÙ„ádájíríyíð®ðÐðéðöó¢òöòõòæòâòÕòÊòÄò¾ò£ñÁóîóÛóÄó¿ñ¿ñ¾ñºñ´ñ¯ñ«ñ¨ðùðõðêðàðÖðÎðÁð°';
  v_Charz   Varchar2(2000) := 'Ø´ØÆØÓØëØùÙªÙ¾ÙÌÙÞÚ£Ú¯ÚºÚÁÚÂÚÑÚØÚÚÚÞÚèÚìÚîÛ¤Û¥ÛµÛ¸ÛÚÛúÜÆÜÑÜïÝ§ÝÏÝèÞ©ÞÊÞÙÞêÞøÞýß¡ß¤ßªß¬ß¸ßÆßåßîßðßòßõßùßúàùàýá¤á¿áÌáÑáÖáçâ¯âÍâåâôã·ä¥ä¨ä·ä¾äÃäóåªåÅåéåëæ¢æ¨æÑæÜæàæãæíæïæûç§ç»çÄçÇçÕçÚçÞè°è¶èÌèÎèÏè×èÙèäèåè÷èþé«é»éÆéÍé×éòéôéùéüê¢ê°êµê¾êÃêÞêßêâëÆëÐëÑëÓëÕëÞëêëùì¹ìÄìíìñìóìõìúí§í½íÄíÎíØíÝíéíöî³îÀîÈîÛîùï£ï­ïÅïßïíïñïôïöð¡ð²ðºðÑðäðæðëðññ©ñ¸ñÞñèò§òÆòÎòØó®ó¯óÃóÉóÊóÐóÝóåóçóðô¢ô¦ô±ô¶ô·ôÒôÕôØôãôêôõô÷ôüõ¥õÅõÙõÜõàõîõòõôõþö£ö¤ö·öÉöíöö÷®÷Ú÷æ÷þ';

Begin
  If v_OutNum<1 Or v_OutNum>40 Then
     v_OutMaxNum:=10;
  Else
    v_OutMaxNum:=v_OutNum;
  End If;

  If v_Instr Is Null Or Length(Ltrim(v_Instr)) = 0 Then
    v_Spell := '';
  Else
    v_Input := Upper(v_Instr);
    v_Spell := '';
    For v_Bitnum In 1 .. Length(v_Input) Loop
      v_Bitchar := Substr(v_Input, v_Bitnum, 1);
      If v_Bitchar >= '°¡' And v_Bitchar <= '×ù' Then
        For v_Chrnum In 1 .. Length(v_Stdstr) Loop
          If Substr(v_Stdstr, v_Chrnum, 1) = '-' Then
            Null;
          Elsif v_Bitchar < Substr(v_Stdstr, v_Chrnum, 1) Then
            v_Spell := v_Spell || Chr(64 + v_Chrnum);
            Exit;
          End If;
        End Loop;
        If v_Bitchar >= 'ÔÑ' Then
          v_Spell := v_Spell || 'Z';
        End If;
      Elsif Instr('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+-*/', v_Bitchar) > 0 Then
        v_Spell := v_Spell || v_Bitchar;
      Elsif Instr('¢ñ¢ò¢ó¢ô¢õ¢ö¢÷¢ø¢ù', v_Bitchar) > 0 Then
        v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41664);
      Elsif Instr('£Á£Â£Ã£Ä£Å£Æ£Ç£È£É£Ê£Ë£Ì£Í£Î£Ï£Ð£Ñ£Ò£Ó£Ô£Õ£Ö£×£Ø£Ù£Ú',v_Bitchar) > 0 Then
        v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41856);
      Elsif Instr('¦¡¦Á', v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'A';
      Elsif Instr('¦¢¦Â', v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'B';
      Elsif Instr('¦£¦Ã', v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'G';
      Elsif Instr(v_Chara, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'A';
      Elsif Instr(v_Charb, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'B';
      Elsif Instr(v_Charc, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'C';
      Elsif Instr(v_Chard, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'D';
      Elsif Instr(v_Chare, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'E';
      Elsif Instr(v_Charf, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'F';
      Elsif Instr(v_Charg, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'G';
      Elsif Instr(v_Charh, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'H';
      Elsif Instr(v_Charj, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'J';
      Elsif Instr(v_Chark, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'K';
      Elsif Instr(v_Charl, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'L';
      Elsif Instr(v_Charm, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'M';
      Elsif Instr(v_Charn, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'N';
      Elsif Instr(v_Charo, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'O';
      Elsif Instr(v_Charp, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'P';
      Elsif Instr(v_Charq, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'Q';
      Elsif Instr(v_Charr, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'R';
      Elsif Instr(v_Chars, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'S';
      Elsif Instr(v_Chart, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'T';
      Elsif Instr(v_Charw, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'W';
      Elsif Instr(v_Charx, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'X';
      Elsif Instr(v_Chary, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'Y';
      Elsif Instr(v_Charz, v_Bitchar) > 0 Then
        v_Spell := v_Spell || 'Z';
--      Else
--        v_Spell := v_Spell || '_';
      End If;
      Exit When Length(v_Spell) > v_OutMaxNum-1;
    End Loop;
  End If;
  Return(v_Spell);
End;
/