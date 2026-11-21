/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where Windows Word hyphenates American English words that the TeX patterns
 * (./hyphenation-en-us.ts) break elsewhere, or not at all, measured in Word
 * (brk12/probes/mkw.py: each word alone on a line just wide enough for its
 * first k letters and a hyphen, k = 2 .. length - 2, Calibri 11, mode 15,
 * autoHyphenation; Word broke at k exactly when k is one of its points, and
 * every earlier break it fell back to was one of them too). 9,594 words
 * measured: the corpus's commonest and the TUGboat US exception list's;
 * 2,693 differ from the patterns and are listed here, hyphenated
 * at Word's points ("anal-y-sis", "nat-u-ral", "pro-cess", "alt-hough"; a
 * word without a hyphen Word does not break: "knowledge", "process" is
 * "pro-cess"). Read by ./hyphenation.ts over the patterns' own exceptions.
 */
export const WORD_EN_US_EXCEPTIONS =
  "aadhaar ab-dulmaliq ab-rupt ab-so-lute ab-surd-ity abu-ba-kar ac-a-demic ac-a-dem-ics ac-cepta-ble ac-ceptance " +
  "ac-cord-ance ac-count-a-bil-ity ac-count-a-ble ac-cu-sa-tive ac-ids acknowledge acknowl-edged acknowl-edge-ment " +
  "acknowl-edge-ments acknowl-edges acous-tic ac-ro-nym ac-ro-nyms acrylal-de-hyde acryla-mide acryla-mides " +
  "ac-tiv-ism ac-tualmente ac-u-punc-ture ac-u-punc-tur-ist ad-a-mant ad-ams adapt-a-bil-ity adapt-a-ble adaptarse " +
  "ad-ap-ta-tion ad-ap-ta-tions ad-dable addi-ble add-ing adebayo ad-e-lante ad-min-istra-tion ad-min-istra-tions " +
  "ad-o-les-cents adustus ad-vantage ad-vantages ad-vent-ist aer-o-space af-ford-a-bil-ity af-ghan-i-stan afolabi " +
  "af-rica af-ri-can ag-ri-cul-tural ag-ri-cul-ture ag-ro-nomic ah-med akinbote akinola al-drich al-ex-an-der " +
  "al-ex-an-dria al-ex-an-drine al-ge-bra-i-cally al-ge-bra-ische al-gon-quian al-gon-quin al-ien-a-tion " +
  "al-i-mo-sho al-le-gheny al-len al-li-ance al-ma-nac al-ma-nack al-ma-nacks al-ma-nacs al-pinus al-ter-a-tion " +
  "al-ter-a-tions alt-hough ama-lar am-a-zon am-biv-a-lent ame-nemope amindar-bari am-phip-oda am-phi-pods " +
  "am-ster-dam anach-ro-nism anach-ro-nis-tic anah-tar an-a-lect an-a-lects an-al-ge-sics an-a-liz an-a-lizar " +
  "an-a-lizi an-a-logue an-a-lyse an-a-lysed anal-yses an-a-lys-ing anal-y-sis an-a-sta-si-adou anat-omy " +
  "an-da-lu-cia an-da-lu-cian an-da-lu-sia an-da-lusian an-eu-rysm an-eu-rys-mal an-eu-rysms an-i-so-tropic " +
  "an-i-so-trop-i-cally an-i-so-tropism an-i-sot-ropy an-ni-ver-saries an-ni-ver-sary anom-a-lies " +
  "anom-a-lop-terygella an-o-nym-ity anon-y-mous anon-y-mously ansberry an-tes an-ti-bi-otic an-ti-bi-ot-ics " +
  "an-ti-hol-o-mor-phic an-ti-mi-totic an-tinomy an-ti-thet-ical an-ti-vaxx-ers anwar an-y-one an-y-thing " +
  "an-zei-ger apa-rici apge-o-chem apgeog aplicada apol-lodo-rus ap-pli-ca-ble aprox-i-mad-amente aprutiensis " +
  "aquat-icus aquellas ar-a-bic ar-a-maic ar-chi-me-dean ar-chi-pel-ago ar-chi-pel-a-gos ar-chive ar-chives " +
  "ar-chiv-ing ar-chi-vist ar-chi-vists arctan-gent arctan-gents ar-eal ar-e-tleri ar-gen-tina arith-me-tic " +
  "ar-kan-sas ar-range ar-range-ment ar-range-ments ar-ri-val ar-se-nic ar-te-facts ar-tery ar-thri-tis art-ist " +
  "art-istry art-ists asegurar as-sessing as-sign-ors as-sis-tantship as-sis-tantships as-so-ci-ate as-the-non " +
  "as-trol-o-ger as-trol-o-gers as-tron-o-mer as-tron-o-mers ath-ens atpase atpases at-tribut-a-ble at-trib-ute " +
  "at-tributed at-trib-uting at-trib-u-tion audzi-jonyte auflage aus-tral-a-sian aus-tral-ian au-thor-i-za-tion " +
  "au-thor-ize au-tom-ata au-to-mat-ing au-to-ma-tion au-tom-a-tis-ierter au-tom-a-ton au-ton-o-mous " +
  "au-tonumber-ing au-to-regres-sion au-to-regres-sive au-to-res au-to-round-ing avail-a-bil-ity avail-a-ble " +
  "av-a-tar av-e-nues ayur-veda ba-ba-lola ba-ba-lo-laso bac-chi-o-chi backend back-pedal back-ped-al-ing " +
  "back-ped-als back-scratcher back-scratch-ing ba-ha-mas bah-rain ba-lu-chi-stan ban-dleader ban-dlead-ers " +
  "bank-roll bank-rolls bario-phar-ma-co-sider-ite bar-o-nies base-lineskip ba-ses basic bas-sou-lis bas-tien " +
  "bath-y-scaphe bberstedt beb-chuk be-dienung be-drag-gle be-drag-gled bed-rid-den bed-rock bee-tho-ven begin " +
  "beklen-mekte-dir bel-a-rus be-lir-len-mesine bembo bender be-nev-o-lence benzler ber-lin ber-nado ber-nardo " +
  "bev-ies bey-ene bi-ble bib-li-o-graph-ical bib-li-og-ra-phische bib-li-og-raphy bib-li-og-raphystyle bib-un-its " +
  "bidif-fer-en-tial bi-di-rec-tional bi-ge-low big-gest bi-gut-ta-tus bild-irilen bild-irilmi bi-limsel bill-a-ble " +
  "bi-o-ac-ces-si-bil-ity bi-o-a-vail-a-bil-ity bi-o-chem-i-cal bi-o-chem-is-try bi-o-di-ver-sity bi-o-ge-og-ra-phy " +
  "bi-o-in-for-mat-ics bi-o-mass bi-o-math-e-mat-ics bi-o-mat-ters bi-o-med-i-cal bi-o-med-i-cine bi-o-metrika " +
  "bi-o-re-sources bi-o-rhythms bi-ota bi-o-tech-nol-ogy bi-o-tic bio-weaponry bio-weap-ons bi-po-lar birlikte " +
  "bir-ming-ham bit-map bit-maps bi-va-ri-ate biy-olojik block-chain bod-dupalli bo-he-mian bol-i-vari bo-lom-e-ter " +
  "bo-lom-e-ters bol-yai bookseller booksellers bool-ean bool-e-ans bos-kovici bos-nia bos-ton botosa-ne-anui " +
  "bound-a-ries bra-sil bra-zil bra-zil-ian bring-hurst brink-mann brit-ish brodhun bruns-wick brus-quer bu-da-pest " +
  "budget bueh-ler bue-nos buff-ers bul-le-tin bulunan bulundu bulun-mak-tad bulun-mamak-tad bun-desamt bun-dling " +
  "burck-hardt burn-out bus-ier bus-i-est bus-sing buzzword buzzwords cache-abil-ity cache-able ca-dre call-ers " +
  "camada cam-er-a-men can-ada ca-na-dian ca-nal ca-non-i-cal ca-pac-i-dad cap-ita ca-rac-ter carathe-o-dory " +
  "car-di-o-vas-cu-lar care-giv-ers car-ib-bean car-min-ite car-ne-gie car-pa-thi-cus cas-pian cas-u-alty " +
  "cat-a-lysed catarrhs ca-tas-tro-phe ca-tas-tro-phes catas-tro-phism cat-e-go-ries cat-e-go-ri-sa-tion " +
  "cat-e-go-ri-za-tion cat-e-noid cat-e-noids ca-veat ce-lina cengiz cen-trif-u-gal cen-tu-ries cer-co-spora " +
  "cer-e-bral cer-tif-i-cate cer-tif-i-cates cha-me-leon cha-me-le-ons chan-cery change-a-ble cha-otic " +
  "char-ac-ter-ised charles-ton char-met chem-is-try chemo-au-to-trophic chem-o-kine chem-o-kines chemo-ther-a-pies " +
  "chem-o-ther-apy ches-ter chhattis-garh chiang chi-cago chich-es-ter chidiebere chloro-methane chloro-methanes " +
  "chorlim chris-tian chro-ma-to-grams chron-o-log-i-cal chrys-otus cit-ies ci-vil-ian civ-i-li-za-tion " +
  "clas-ifi-caci clas-ses cli-ni-cian cli-ni-cilor clon-ing co-a-li-tions coasso-ci-a-tive coch-lea coch-lear " +
  "coch-leas codesigner codesign-ers co-efi-ciente co-ex-ist-ence cog-nac cog-nacs co-gno-scenti co-her-en-cia " +
  "coker-nel coker-nels col-lin-e-a-tion co-lom-bia co-lo-nial co-lo-ni-al-ism col-o-ni-za-tion col-o-phon " +
  "col-o-phons col-ours co-lum-bian col-um-bine col-um-bines co-lum-nar col-umns comer-cio com-fort-a-ble " +
  "com-mod-i-ties com-muni-cate com-munist com-parai-son com-parands com-par-a-tor com-par-i-sons com-pen-dium " +
  "com-pet-i-tors com-ple-men-ta-ri-ties com-po-nen-tes com-po-nentwise com-po-site com-put-a-bil-ity com-put-a-ble " +
  "com-pu-ting co-mun-i-dad con-cord-ance con-du-cive con-fi-den-ti-al-ity con-fig-u-ral con-firm-a-tory " +
  "con-form-a-ble con-form-ist con-form-ists con-form-ity con-fron-ta-tion con-fron-ta-tional con-front-ing " +
  "con-ge-ries conoci-miento conoci-mien-tos con-serva-tive con-sid-er-a-ble con-sider-ando con-sistency " +
  "con-sistent con-sist-ently con-stit-u-ents con-structi-bil-ity con-structi-ble con-sum-ers con-tam-i-na-tion " +
  "con-tig-u-ous con-trac-ta-ble con-trasting con-trib-ute con-trib-uted con-trib-utes con-trib-uting con-ven-ient " +
  "co-per-ni-cus cop-ies cop-y-right cop-y-right-a-ble co-re-li-gion-ist co-re-li-gion-ists co-re-op-sis " +
  "corespond-ent corespond-ents cor-rel-aci cor-re-spond-ence cor-re-spond-ences cor-re-spond-ents " +
  "cor-ti-co-ster-oid co-sas co-se-cant co-semi-sim-ple cos-mo-chim co-tan-gent coun-sel-ling coun-ter cour-age " +
  "cou-ra-geous cou-ra-geously cov-e-nant co-vers cran-gonycti-dae crank-case cre-a-tion cre-a-tive " +
  "cre-a-tivecom-mons cre-a-tiv-ity cre-a-tor crec-i-miento crenal cri-ses cro-a-tia croati-cus cronbach " +
  "cross-o-ver crus-ta-cea crus-ta-cean crus-ta-ceans cryp-to-gram cryp-to-grams cu-ando cu-atro cui-sines " +
  "cu-nei-form cu-ri-os-ity cur-ric-u-lum cus-tom-ers cus-tom-i-za-ble cus-tom-ize cus-tom-ized cus-tom-izes cyber " +
  "cybervirus cyberviruses cyberweapon cyberweap-ons cy-prus cy-to-chrome cy-to-kine cy-to-kines czech-o-slo-vakia " +
  "dac-tyl-o-gram dac-tylo-graph dalecar-lica da-mar dam-sel-flies dam-sel-fly dan-iel da-ta-base da-ta-bases " +
  "da-ta-flow datapath datapaths data-point data-points da-taset da-tasets dates-tamp dates-tamps datetime da-tos " +
  "da-vid da-vis dav-ran debido dec-ade dec-ades de-clar-a-ble dec-la-ra-tion dec-la-ra-tions de-fin-ing " +
  "def-i-ni-tion def-i-ni-tions deg-ra-da-tion de-ion-ized de-i-ties del-a-ware de-lin-e-ated dem-i-sem-i-qua-ver " +
  "dem-i-sem-i-qua-vers de-moc-racy dem-o-crat dem-o-cratic de-mocratism dem-o-crats de-mographics de-mon-stra-ble " +
  "de-mon-stra-bly de-mos de-mostrando de-part-mental de-pend-ence de-pend-en-cies de-pend-ency de-pend-ent " +
  "de-pos-ita de-pos-ited dep-o-si-tion de-pos-its dep-uty der-inleme-sine der-i-va-tion der-i-va-tional " +
  "der-i-va-tions de-scend-ants de-sert de-sir-a-ble despu de-sta-bi-lize destacando de-tect-a-ble devlet " +
  "di-a-be-tes di-ag-no-ses di-a-lect di-a-lec-tal di-a-lec-tic di-a-lec-ti-cal di-a-lec-ti-cian di-a-lec-ti-cians " +
  "di-a-lec-tics di-a-lects di-a-met-rical di-a-tom di-a-to-ma-ceous di-a-toms di-chloro-methane di-chot-omy " +
  "dic-tion-ar-ies dic-tion-ary di-ego diferente dig-i-tal-i-za-tion dijks-tra dik-kat di-lemma di-lem-mas " +
  "di-og-e-nous di-ox-ide dip-loid di-plo-macy dip-lo-matic di-rec-tamente di-rective dis-a-bil-ity dis-aggre-gated " +
  "dis-a-gree dis-a-gree-ment dis-cri-mi-nant dis-cuss-ing dis-mis-sal dis-par-and dis-par-ands dis-poni-bles " +
  "dis-trib-ut-a-ble dis-trib-uting dis-tri-bu-tion dis-tri-bu-tions dis-turb-ance dis-turb-ances dis-u-til-ity " +
  "do-ble do-min-i-can dor-ches-ter dou-ady dou-ble-spac-ing dou-ble-talk drahota drechsler driven driver driv-ers " +
  "drives driv-ing drom-e-dar-ies drop-let drop-lets drop-out dro-soph-ila dru-sus duane du-choslav du-o-po-lies " +
  "du-o-po-list du-o-po-lists du-o-poly du-rumlar duvick dwell-ing dy-nam-ical dy-na-mische dys-lec-tic dys-lexia " +
  "dysp-noea dysreg-u-la-tion dys-to-pia eastend-ers ec-cli-sop-te-ryx econ-o-mies eco-sys-tem eco-sys-tems " +
  "ec-ua-dor ederek edile-bilir ed-ited ed-it-ing ed-i-to-rial ed-u-ca-tivo ed-u-ca-tivos eg-gler egg-shell " +
  "egg-shells egyp-tian ei-gen-class ei-gen-classes ei-gen-value ei-gen-val-ues eijk-hout ekilde ekiti " +
  "elec-tro-en-ceph-a-lo-gram elec-tro-lyte elec-tro-me-chan-i-cal elec-tro-mech-a-no-acous-tic elec-tro-pho-re-sis " +
  "elec-tro-pho-retic el-ektrik el-ektronik ele-mentos ellos elonmusk else-vier eman-uel en-a-ble en-a-bled " +
  "en-a-bles en-a-bling en-coun-ter en-de-mism end-game en-do-crine en-do-cri-nol-ogy en-dog-e-nous en-dos-co-pies " +
  "en-dos-copy engle englisch en-tender en-tor-nos en-vi-ron ephem-era eph-e-mer-i-des ephem-eris ep-i-demic " +
  "ep-i-de-mi-o-log-i-cal ep-i-de-mi-ol-ogy ep-i-graphs ep-is-temic equi-di-men-sional eq-ui-ta-ble equivar-i-ance " +
  "equivar-i-ant er-go-nomic er-go-nom-i-cally er-go-nom-ics er-lendir-mel-erde erle-rinin erucic esen-cial " +
  "esfuer-zos es-sence es-sences esta-bi-lidad esta-dis-ti-cas estruc-tura es-tu-diantes eth-ics ethyl-a-mine " +
  "eth-ylene ethynyl-a-tion etkiler etkileri etkile-rini etkile-rinin etkili etkisi etkisini etkisinin et-mekte-dir " +
  "eu-lerian eu-sta-chian ev-ans-ton even-ing eve-ry-day eve-ry-one eve-ry-thing evresel evreye ex-e-get-i-cal " +
  "ex-e-get-i-cally ex-hib-its ex-ist-ence ex-og-e-nous ex-pendi-ture ex-per-i-mental ex-plan-a-tory " +
  "ex-ploi-ta-tion ex-plo-ra-tion ex-plo-ra-tions ex-plor-a-tory ex-sero-hi-lum ex-tracta-bil-ity ex-traor-di-nary " +
  "fa-ce-book fa-ko-rede fall-ing fa-mil-i-ares fa-mil-iar-i-sa-tion farma-kolojik fa-vor-a-ble fa-vor-a-bly " +
  "fa-vour fa-vours fazla fer-mi-ons fer-ti-lizer fest-schrift fe-ver ffen-tlichen fi-a-bilidade fi-du-ci-ar-ies " +
  "fi-du-ci-ary fighting fig-u-rine fig-u-rines filippi fi-nal-ise fintech flor-ida flo-rid-ian flowchart " +
  "flowcharts fluoros-co-pies fluor-os-copy fon-tenc for-ests for-mation form-a-tive for-schung-sin-sti-tut " +
  "for-ster fra-gil-ity frank-lin free-sia free-sias frie-drich friend-li-est fri-vol-i-ties fri-vol-ity frontend " +
  "frontends fro-zen fuer-zas fun-cion-aba funktsional ga-lac-tic gal-ax-ies gal-axy gam-mar-i-dae gaz-et-teer " +
  "gaz-et-teers gelecek gelecekteki ge-le-neksel ge-ne-a-log-i-cal ge-ne-ious genel-li-kle gen-er-aci gen-er-al-ise " +
  "gen-er-al-ised gen-er-al-i-za-tion gen-er-al-i-za-tions gen-erar ge-neric ge-net-ically ge-nome ge-nomewide " +
  "gen-o-type gen-o-types gen-o-typic gen-u-ine gen-u-inely ge-nus ge-o-chem ge-o-chem-i-cal ge-o-chem-is-try " +
  "ge-o-chim ge-oestad geof-frey geofond ge-o-ge-bra ge-o-graf ge-o-graph-ical ge-om-e-ter ge-om-e-ters " +
  "ge-o-met-ric ge-omet-rics ge-om-e-try ge-o-po-lit-i-cal ge-o-sta-tis-ti-cal ge-o-strophic ge-o-ther-mal ger-ekti " +
  "ge-sell-schaft ghost-view gi-gan-odes gil-more gitee glob-al-i-za-tion gott-fried gottlieb gov-ern-ance " +
  "gran-di-ose grand-un-cle grand-un-cles graph-eme graph-emes gra-phemic gra-phem-ics grapho-lin-guis-tic " +
  "grapho-lin-guis-tics greifswald grothend-ieck grouplike grundleh-ren gruy-ter grygar gui-ana ha-ber hab-i-tat " +
  "hab-its hada-mard had-ith half-tone half-tones half-way hallauer ham-il-ton ham-il-to-nian har-mo-ni-ous " +
  "har-mo-ni-za-tion has-tal hatch-er-ies ha-ven haz-ard-ous healthcare he-ge-lian he-gem-ony hei-nous " +
  "he-li-o-pause he-li-o-trope hel-lin hel-sinki hem-i-dem-i-sem-i-qua-ver hem-i-dem-i-sem-i-qua-vers " +
  "hem-orhe-ol-ogy her-findahl her-it-a-bil-ity her-it-a-ble her-it-age hervant her-ze-govina het-er-o-ge-ne-ity " +
  "het-er-o-ge-ne-ous het-ero-inter-face het-ero-inter-faces het-er-o-sis heu-ris-tic hex-a-dec-i-mal hez-e-kian " +
  "hibbs hi-er-ar-chical hil-bert hipo-tecaria hip-po-pot-a-mus hirsch-man his-tone hoefler hoekwater hok-kaido " +
  "ho-lis-tic hol-o-caust ho-lodeck ho-lodecks ho-lonom ho-lon-o-mies ho-lon-omy ho-lub ho-me-o-morphic " +
  "ho-me-o-morphism ho-me-o-sta-sis ho-me-o-static ho-me-o-stat-ics hom-i-cide ho-mo-ge-ne-ity ho-mo-ge-ne-ous " +
  "ho-mog-e-nous hom-o-nym hom-o-nymic ho-mon-y-mous hom-o-nyms ho-mon-ymy ho-moph-o-nous ho-moph-ony " +
  "ho-mo-zy-gosity ho-mo-zy-gous ho-nice hon-our ho-ri-zons hos-pi-tal-ised hot-bed hot-beds hou-zar hussain " +
  "hy-brid-i-sa-tion hy-brid-i-za-tion hy-dro-ther-mal hy-perledger hy-perpa-ram-e-ters hy-phen-a-tion " +
  "hy-phen-a-tions hy-po-elas-tic-ity hy-poth-e-ses iba-dan ico-no-grapher ico-no-graphers icon-o-graphic " +
  "ico-no-gra-phy ide-als iden-ti-fi-a-ble iden-ti-fi-ers ide-o-graphs ide-o-lo-gies id-i-o-lect id-i-o-lects " +
  "id-i-o-syn-cra-sies id-i-o-syn-crasy id-i-o-syn-cratic id-i-o-syn-crat-i-cally ig-norespaces ih-tiya ikor-odu " +
  "ilek-troniko ilgili il-li-nois il-liq-uid il-liquid-ity im-age-mag-ick im-ag-i-nary im-ag-ine im-ag-ing " +
  "im-mun-iza-tion im-muno-mod-u-la-tory im-pe-rial im-portance im-portant im-provisus in-a-bil-ity inagaki " +
  "incelemi incelen-mesi incelenmi inceley-erek in-ci-dental in-con-sist-en-cies in-de-pend-ence in-de-pend-ent " +
  "in-de-pendently in-di-ces in-dig-e-nous in-dis-pen-sa-ble in-do-ne-sia in-du-bi-ta-ble in-dus-triales inegi " +
  "in-e-qual-ity in-eq-ui-ta-ble in-ev-i-ta-ble in-flam-atuar in-flu-en-ciado in-flu-enza in-for-mation " +
  "in-forma-tive in-gridae in-her-itance in-hib-i-tory in-i-tial in-i-tial-i-za-tion in-i-tially in-i-ti-ated " +
  "in-i-ti-ates in-i-ti-at-ing in-i-ti-a-tion in-i-ti-a-tive in-i-ti-a-tives in-i-ti-a-tors in-ju-ries " +
  "in-no-va-tive in-putenc in-sig-nif-i-cant in-stall-ers in-stitu-ci-ones in-sti-tu-tion-al-ised " +
  "in-sti-tu-tion-al-i-za-tion in-teg-rity in-ter-de-pend-ence in-te-rele-ment in-ter-ga-lac-tic in-ter-na-tio-nale " +
  "in-terop-er-a-bil-ity in-terop-er-a-ble in-ter-pret-a-bil-ity in-ter-pret-a-ble in-ter-pret-ers in-ter-quar-tile " +
  "in-ter-viewee in-ter-view-ees in-trac-ta-bil-ity in-trac-ta-ble in-tu-i-tively in-u-tile in-u-til-ity " +
  "in-var-i-ance in-var-i-ant ira-nian ir-re-duc-i-ble ir-re-duc-i-bly isaiah iso-ge-omet-rics isop-oda " +
  "iso-ther-mal iso-tropic isot-ropy it-aly itin-er-ar-ies iyile jackow-ski ja-leta ja-natka ja-pan jap-a-nese " +
  "ja-vas-cript jer-e-miad jer-e-mi-ads jer-e-miah jhar-khand jinping jis-uan jo-seph julia ju-lie jung-ian jun-ior " +
  "justus ka-blo ka-blolar ka-bul ka-domtsev ka-kamega kal-enda kal-mak-tad kan-sas ka-pa-sitesi kaplan kapsam " +
  "kapsaml ka-ren karls-ruhe kar-na-taka ka-sha kat-e-go-ris-inde ka-zakh-stan kelimeler ken-dall keynes-ian " +
  "key-note key-notes khai-ral-lah kilifi kill-ing ki-lonodes kitalale kitale klementov knowledge knowl-edge-a-ble " +
  "koeh-ler kol-mo-go-rov kontrol kopter-ides korteweg koruma ko-runmas koruyucu ko-ymak-tad ko-ymu kraj-inu " +
  "krishna krish-na-ism krish-nan kritik kron-ecker krusniki kuchanur kuemmer-len ku-lakow-ski kunstakad-e-mie " +
  "la-bem la-bor-a-to-ries la-bor-a-to-rio la-bor-a-tory la-bour lab-y-rinth lab-y-rin-thian lab-y-rin-thine " +
  "lab-y-rinths la-mak lan-cas-ter la-place la-rak larg-est lasbela las-ued lat-via la-voie lay-arak la-zi-ness " +
  "leaf-hop-per leaf-hop-pers leaf-let leaf-lets le-bilir leg-a-cies leg-acy leg-ate leg-ates le-ga-tion " +
  "le-ga-tions le-gen-dre leg-is-la-tive lep-i-dop-tera lep-todru-sus ler-inde les-sons let-ter-spac-ing " +
  "leu-ko-cyte leu-ko-cytes leu-ko-tri-ene leu-ko-tri-enes levene lev-er-age lev-er-aged le-vit-i-cal le-vit-i-cus " +
  "lie-big lifespan lifespans life-style life-styles lift-off lim-ita lin-e-age lin-e-ages line-spac-ing linun " +
  "li-on-ess li-quid-ity lisis lit-e-rat lit-er-a-tura lith-u-a-nia li-toch-leb li-toch-lebov live-li-hoods llevar " +
  "lma-lar lo-cal-i-za-tion logics lo-gistic logos lojban long-est lon-gev-ity long-term lookahead lo-quac-ity " +
  "lou-i-si-ana luala-tex lujza lytras mac-mil-lan mac-ro-e-co-nomic mac-ro-e-co-nom-ics mac-ro-e-con-omy " +
  "mac-ro-gen mac-ros mad-ame ma-gel-lan ma-go-ro-ko-sho ma-ha-bha-rata ma-ha-mood mainte-nance majid ma-kale " +
  "ma-ka-lenin makeindex ma-kumbi mal-a-chi ma-laise mal-a-prop-ism mal-a-prop-isms ma-lar ma-lara ma-larda " +
  "mal-a-ya-lam ma-lay-sian malickyi ma-lin-vest-ment ma-lin-vest-ments man-a-ge-rial man-ag-ers man-da-tory manera " +
  "man-kind man-slaugh-ter ma-ria mar-i-time marktoberdorf ma-ruz mash-ingai-dze mas-sa-chu-setts mas-sa-cre ma-tem " +
  "ma-te-rial ma-te-riales math-e-ma-ti-cian math-e-ma-ti-cians ma-thias ma-turity max-i-mil-ian max-im-ize " +
  "max-im-iz-ing max-well maydis meas-ur-a-ble meas-ure meas-ured meas-ure-ment meas-ure-ments measures meas-ur-ing " +
  "me-canis-mos med-i-caid med-i-cal med-i-ca-tion med-i-ca-tions med-ici me-dic-i-nal me-dio me-di-o-cre " +
  "me-di-oc-ri-ties me-dium meg-a-fauna meg-a-fau-nal meg-a-lith meg-a-liths meg-a-luro-pi-dae me-ga-nodes " +
  "mekaniz-ma-lar mekte-dir melanchaetes meli-tidae me-lodic memory men-de-ley men-kir mensch me-rid-ionalis " +
  "mes-lek mes-o-po-ta-mia mes-o-po-ta-mian met-a-bolic me-tab-o-lite me-tab-o-lites metadata met-a-form " +
  "met-a-forms met-a-lan-guage met-a-lan-guages metanoea met-a-phor met-a-phor-i-cal met-a-phor-i-cally met-a-phors " +
  "met-a-sta-bil-ity met-a-sta-ble metatable metatables met-a-zoan me-tem-psy-cho-sis meth-am-phet-a-mine me-thane " +
  "meth-od-ism meth-od-ist meth-od-o-log-i-cal meth-od-o-log-i-cally meth-od-ol-o-gies meth-od-ol-ogy " +
  "me-thylammo-nium meth-yl-ate meth-yl-ated meth-yl-a-tion meth-ylene metropoles mevcut mi-chael mi-cro-da-tos " +
  "mi-cro-e-co-nomic mi-cro-e-co-nom-ics mi-cro-e-con-omy mi-cro-en-ter-prise mi-cro-en-ter-prises " +
  "mi-cro-or-gan-ism mi-cro-or-gan-isms mi-cro-struc-ture midaft-er-noon mihaljevi mi-klos mi-kutta millage mil-ler " +
  "mil-li-li-ter mil-li-li-ters mim-e-o-graphed mim-e-o-graphs min-er-al-i-za-tion mine-sweeper mine-sweep-ers " +
  "min-i-mi-za-tion min-i-prep min-is-te-rial min-is-try min-i-sym-po-sia min-i-sym-po-sium min-kow-ski " +
  "min-ne-ap-olis min-ne-sota mi-nor-i-ties minuter minutes minut-est misattrib-u-tion mis-in-for-mation " +
  "mi-to-chon-drial mi-troo mne-monic mne-mon-ics mo-dal-i-ties mo-dal-ity modelo modelos mo-der-nity " +
  "mod-ern-i-za-tion moderno mohini mo-lec-u-lar mo-lenaar molete momoh mon-aco mon-ey-lender mon-ey-lend-ers " +
  "mo-nocen-tra mon-o-chrome mo-noen-er-getic mo-noid mon-o-layer mon-oph-thong mon-oph-thongs mon-o-pole " +
  "mon-o-poles mon-o-space mon-o-spaced mon-o-spac-ing mon-o-spline mon-o-splines mo-nos-tro-fic mon-ta-gnard " +
  "mon-ta-gnards mo-rales mo-ral-ity mor-ally mo-ran mo-rans mo-ra-vian mo-ravsk mor-dell more-o-ver mor-fol-o-jik " +
  "mos-cow mos-kva mottaleb mou-lin muelleri mu-ham-mad mul-ticlass mul-ti-col-lin-e-ar-ity mul-ti-plica-bly " +
  "mul-ti-pli-ca-tive mul-ti-user mul-ti-var-i-a-ble mul-ti-var-i-ate mun-dial my-sida na-chrichten nacional " +
  "nai-robi namespace namespaces na-poca nash-ville na-ta-raj na-tur nat-u-ral na-turschutz nav-i-gating nceden " +
  "neb-u-licola nece-sidad nece-sidades nec-es-sarily neden nedeni-yle nedenle neg-a-ti-vamente nelik neofield " +
  "neofields ne-o-na-tal ne-on-azi ne-on-a-zis ne-pal neph-ews neph-rite ne-phritic nerv-ous net-bsd neth-er-lands " +
  "neu-ral neu-ro-sci-ence new-est news-let-ter news-let-ters nic-a-ra-gua nich-o-las nic-o-sia nie-tzsche ni-ger " +
  "ni-ge-ria ni-ge-rian ni-ge-ri-ans nij-me-gen ni-phargid ni-phargi-dae ni-phar-gus ni-tro-me-thane ni-ver-si-tesi " +
  "nodel-ist nodel-ists no-ether-ian noname nonar-chi-me-dean nonarith-me-tic nonemer-gency non-e-quivar-i-ance " +
  "none-the-less noneu-clid-ean non-i-so-mor-phic nonpseu-do-com-pact non-u-ni-form non-u-ni-formly non-zero " +
  "noord-wijker-hout nor-dic nore-pi-neph-rine nor-mal-ised nor-mal-i-za-tion north-amp-ton no-sol-ogy no-ta-ble " +
  "no-ta-bly no-tice-a-ble not-with-stand-ing novedoso noviembre nu-cle-o-syn-the-sis nu-cle-o-tide nu-cle-o-tides " +
  "nurse-ries nursery ob-jetivo ob-jetivos ob-li-ga-tion ob-li-ga-tions ob-serv-a-ble ob-stfeld ob-tener " +
  "ocean-og-ra-phy och-onogor oedicerotidae of-fi-ci-ants of-rece oksi-datif olanrewaju olayonuolu-wanifemi " +
  "ol-i-go-nu-cle-o-tide ol-i-gop-o-lies ol-i-gop-o-list ol-i-gop-o-lists ol-i-gop-oly ol-i-ver olmad olmak olmakla " +
  "olma-yan ologunde ol-sen ol-u-dele olu-wanifemi om-ni-pres-ence om-ni-pres-ent online on-o-mat-o-poeia " +
  "on-o-mat-o-po-etic on-tol-o-gies onwuekwe onyiuke op-er-and op-er-ands op-er-a-tion-al-ized opor-tun-i-dades " +
  "op-por-tunity op-ti-mi-za-tion orangu-tan orangutans ore-o-pou-los or-gan-i-za-tion or-gan-i-za-tional " +
  "or-gan-i-za-tions or-gan-ize or-gan-iz-ing origin or-i-gins or-tamlar ortaya or-thokeratol-ogy " +
  "or-thon-i-tro-tol-u-ene oshi-mili oshin out-li-ers over-em-pha-sised over-ex-tended over-treat-ment over-view " +
  "over-views ox-a-late ox-i-dic ox-y-gen ox-y-hydr pack-ard pad-ding pag-erank pain-ful pa-ki-stan pa-ki-stani " +
  "pala-bra pa-lermo pal-ette pal-ettes pal-li-dus pam-ela pan-ama papel pa-pua pa-rab-ola pa-rab-o-las par-a-chute " +
  "par-a-chutes paradi-me-thylben-zene par-a-dox par-a-fluorotol-u-ene par-a-graph par-a-grapher par-a-graphs " +
  "par-a-le-gal par-al-lel-ism par-a-mag-netism par-a-medic par-a-med-ics pa-ram-e-ter pa-ram-e-ters " +
  "para-methyl-an-i-sole par-a-metri-za-tion par-a-metrize par-a-mil-i-tary par-a-mount para-mysis pa-ren-tal " +
  "pa-ren-the-ses pa-reto pass-o-ver pas-teurova pa-tent path-o-gens pa-thol-ogy path-o-phys-i-ol-ogy pat-rick " +
  "pat-ron-age ped-a-gog-ies pe-di-at-rics ped-i-gree pe-dosphere pen-al-ties pen-alty pen-in-sula pen-sam-iento " +
  "pen-ta-gons peraza per-cen-tile per-en-nial per-forma-tive per-i-spo-mena per-i-spo-me-non per-me-a-bil-ity " +
  "pesh-a-war petrogr petros pe-trov-ski pfaff-ian phe-nol-phtha-lein phe-nyl-al-a-nine phil-a-del-phia " +
  "phil-ip-pines phi-los-o-pher phi-los-o-phers phi-los-o-phies phi-los-o-phische pho-neme pho-nemes pho-ne-mic " +
  "pho-to-graphs phtha-lamic phthalate phthisis phy-log-eny phy-si-cian phy-si-cians phys-ics pi-ety pipe-line " +
  "pipe-lines pipe-lin-ing plac-a-ble planthop-per planthop-pers plat-eau plat-eaus pleas-ure plugin plugins " +
  "plumbojar-o-site plu-ral plu-ral-istic plu-ri-sub-har-monic pod-lich poign-ant po-land po-licen-tral-ity " +
  "pol-i-cies pol-lacsek pol-len pol-y-an-drous pol-y-an-dry pol-y-cen-tral-ity pol-y-cen-tric pol-y-dac-tyl " +
  "pol-y-dac-tyl-lic pol-y-ene pol-y-eth-ylene polygoni-za-tion po-lyg-y-nous po-lyg-yny pol-y-mer-ase " +
  "polymetal-lic po-lynya po-lyn-yas pol-y-pep-tide pol-y-pep-tides pol-y-phonic po-lyph-ony pol-y-sty-rene " +
  "pol-y-vi-nyl pom-e-gran-ate pom-e-gran-ates po-povi po-ro-e-las-tic po-rous port-a-ble por-tu-guese pos-its " +
  "post-ex-ilic post-hu-mous po-ta-ble potansiyel potansiyeli potansiye-lini po-tato po-ta-toes " +
  "po-ten-tialglei-chung po-to-mac pov-erty pra-gue pred-a-tory pre-dict-a-ble pre-di-lec-tion pre-di-lec-tions " +
  "pre-fers prem-a-ture prem-ise prem-ises prep-a-ra-tion pre-par-ed-ness pre-par-ing pre-pend pre-pended " +
  "pre-pend-ing pre-pends prep-o-si-tion prep-o-si-tional prep-o-si-tions pre-print pre-prints pre-pro-cess " +
  "pre-pro-cessing pre-pro-ces-sor pre-pro-ces-sors pres-by-ter-i-ans pre-school pre-school-ers pre-schools " +
  "pre-sent presenta presen-ta-tion presen-ta-tions presente pre-senter pre-sent-ers pres-ently pre-sents " +
  "pre-sum-a-ble pre-sum-a-bly pret-ty-printer pret-ty-print-ing prev-a-lence prev-a-lent pre-vi-sic pre-wrap " +
  "pre-wrapped pri-mer pri-mers prin-cipio pri-or-i-tise pro-ac-tive pro-ac-tively prob-a-bil-is-tic pro-cesos " +
  "pro-cess pro-cessed pro-cesses pro-cessing pro-ces-sor proc-la-ma-tion produc-tiv-ity profe-sional proge-nies " +
  "pro-gramas pro-gress pro-hib-its pro-ject pro-jects pro-kar-y-ote pro-kar-y-otes pro-kar-y-otic prom-i-nent " +
  "prom-i-nently pro-mis-cu-ous prom-ise prom-ised prom-ises prom-is-ing prom-is-sory prop-a-ganda pro-pel-ler " +
  "pro-pel-lers pro-pel-ling pro-pen-sity properly pro-phetic pro-pia pro-posals pro-pri-e-tary pro-spec-tive " +
  "pro-spects pros-ta-glan-din pros-ta-glan-dins pro-style pro-styles pro-to-lan-guage pro-to-ty-pal prov-e-nance " +
  "prov-erbs pro-vider pro-vid-ers provid-ing prov-ince prov-inces pro-vin-cial pro-virus pro-viruses prow-ess " +
  "proy-ecto proy-ec-tos pruebas pseu-do-code pseu-do-differ-en-tial pseudofi-nite pseudofi-nitely pseudoforces " +
  "pseudographer pseu-donumber pseu-donumbers pseudoword pseudowords psy-chi-at-ric pu-dilov pu-pils " +
  "pu-ri-fi-ca-tion pur-pos-ive put-ting pyong-yang py-thag-o-ras py-thag-o-rean py-thon quad-ratic quad-rat-ics " +
  "quad-ri-lat-eral quad-ri-lat-erals quad-ri-ple-gic quad-ru-ped quad-ru-peds quad-ru-pole quad-ru-poles " +
  "quan-ti-fi-ca-tion qua-siequiv-a-lence qua-siequiv-a-lences qua-siequiv-a-lent qua-si-hy-ponor-mal " +
  "qua-sirad-i-cal qua-siresid-ual qua-sis-mooth qua-sistation-ary qua-sitopos qua-sitri-an-gu-lar qwerty ra-cial " +
  "rac-ism ra-dha-krish-nan ra-di-o-graphic ra-di-og-raphy rag-ged ra-ja-ram rakshit rap-idly ra-rak ra-shid " +
  "ra-ther rathskel-ler ra-tional ra-tionale ra-tion-ales rav-enous ra-viku-mar read-i-ness re-al-i-za-tion " +
  "re-a-nal-y-sis re-ar-range re-ar-ranged re-ar-range-ment re-ar-range-ments re-ar-ranges rea-son-a-ble re-as-sert " +
  "re-cip-ro-cal rec-i-proc-ity rec-og-nise rec-og-nised rec-og-nis-ing rec-on-cil-a-ble rec-ord rec-orded " +
  "re-cord-ing re-cord-ings rec-ords rec-re-a-tional re-cy-cla-ble re-dac-tion re-di-rect re-di-rec-tion " +
  "re-dis-trib-ut-a-ble re-dis-tri-bu-tion re-dis-tri-bu-tional re-duc-i-ble reecho re-fers ref-u-gee ref-u-gees " +
  "re-fu-gia re-fuse re-gime re-gimes reg-is-try reg-u-lar-i-za-tion reich-lin re-im-ple-ment re-im-ple-men-ta-tion " +
  "re-im-ple-mented re-im-ple-ments rei-nov rel-aci rel-a-cion-a-das rel-a-cio-na-dos rel-a-tiv-istic re-li-a-ble " +
  "re-li-a-bly re-li-ance re-li-gieuses re-mark-a-ble remdesivir rem-e-dies re-new-a-ble re-new-a-bles re-nowned " +
  "re-or-der rep-er-toire re-phrase re-phrased re-phrases rep-li-cate rep-li-cates rep-li-ca-tion rep-li-ca-tions " +
  "re-plies re-port-age re-port-ers re-po-si-tion re-po-si-tions re-pos-i-to-ries re-pos-i-tory rep-re-senta-tive " +
  "rep-re-sent-a-tive-ness rep-re-sent-a-tives re-print re-printed re-prints re-pro-duc-i-bil-ity re-pro-duc-i-ble " +
  "resamples re-search-ers res-er-voir re-sid-ual re-sid-u-als re-sil-i-ence re-sistance re-sistances re-sistant " +
  "re-solv-a-ble res-o-nated re-spond-ent re-spond-ents respuesta re-stor-a-ble res-to-ra-tion re-striction " +
  "re-strictions re-sulta-dos re-sur-gence ret-ro-ac-tively ret-ro-fit ret-ro-fit-ted re-us-a-ble re-use re-used " +
  "rev-e-nue re-vi-tal-ise re-wire re-wrap re-wrapped re-write reyn-olds rhae-tica rhet-o-ric rhe-tor-i-cal " +
  "rhe-tor-i-cally rheu-ma-toid rheu-ma-tol-ogy ricas rich-ard rich-ard-son ri-euwerts rig-our ri-vera rlanmas " +
  "rle-rinin rma-lar rob-ert robes-pierre roch-es-ter ro-dri-guez rog-ers rohovec roll-ing ro-man-icus roof-top " +
  "roof-tops roo-se-vel-tov ropro-tek-tif roun-dish roun-dish-ness roundtable roundtables routledge ru-dan rul-ers " +
  "run-na-ble runtime ru-ral rus-sian ry-dberg sa-ber sa-haran sa-hip sa-hiptir sa-lami sal-ary sa-li-ent " +
  "sal-i-vary sal-ta-tion sal-va-dor sand-o-val sa-pi-en-tia sap-o-nin sari-kaya sar-sa-pa-rilla sat-el-lite " +
  "sat-el-lites sau-er-kraut scal-a-ble sca-lar sca-lars scar-city sceneshifter sceneshift-ing sche-matic " +
  "scholartex school-chil-dren schot-tische schrodinger schwa-bacher schweidnitz schwert sci-totenv scru-tiny " +
  "search-a-ble sec-ond-ary sec-re-tar-iat sec-re-tar-i-ats sed-a-tives sedatu sed-en-tary sejkora sellers semana " +
  "sem-idefi-nite sem-i-di-rect semi-homo-thetic semir-ing semir-ings sem-i-sim-ple sem-i-skilled semistable " +
  "sen-ior sep-tu-a-gint se-rere ser-geant ser-geants seroep-i-de-mi-o-log-i-cal ser-vo-mech-a-nism " +
  "ser-vo-mech-a-nisms ses-qui-pe-da-lian set-ups se-ver-ity sev-tsuk shap-a-ble shape-a-ble shi-kutsa shop-lifter " +
  "shop-lift-ing shoreditch sign-age sig-nal-ling sig-nif-i-ca-tiva sig-nif-i-ca-tivos si-lently simply " +
  "sim-ul-ta-ne-ous sim-ul-ta-ne-ously sin-ga-po-rean sin-ga-po-re-ans sing-ing sin-glespace sin-glespaced " +
  "sin-glespac-ing sistema siste-mas si-units skoup slnunic-ode slo-vakia slo-ve-nia smartphone smol-o-tely so-ber " +
  "so-cial-i-za-tion so-cie-dad so-ci-e-tal so-ci-e-ties so-ci-o-ec-o-nomic sol-o-monic so-lute so-lutes sonder " +
  "so-nu-cunda sorumlu sos-tenible south-amp-ton spe-cial-ised spe-cies spe-cif-i-cally spec-i-fic-ity " +
  "spec-tros-copy spe-laeus sphe-roid sphe-roi-dal sphe-roids spill-o-ver spill-o-vers spinors spir-itual " +
  "spokes-man spokes-per-son spokes-per-sons spra-gue squeam-ish srini-va-san standalone stand-ard " +
  "stand-ard-i-sa-tion stand-ard-i-za-tion stand-ard-ized stand-ards static stat-ute stat-utes stat-u-tory " +
  "stee-ple-chase steph-a-niaolanrewaju ste-re-o-graphic ste-ril-ity sterilmi stermekte-dir sticas stokessche " +
  "stoplist stoplists sto-ry-tell-ing strap-hanger stra-te-gic stra-te-gi-cally strat-e-gies strip-tease strnad " +
  "strong-est struc-tur-ally stud-y-ing sty-lis-tic sub-cat-e-go-ries sub-dif-fer-en-tial sub-node sub-nodes " +
  "subop-ti-mal sub-scrib-ers sub-se-lect sub-se-lected sub-se-lec-tion sub-se-lec-tions sub-se-lects subta-bles " +
  "sub-ter-ra-nean sub-u-nit su-chara su-cha-rov suit-a-ble sum-mar-ies sum-ma-rised su-per-der-i-va-tion " +
  "su-per-der-i-va-tions su-per-ego su-per-e-gos su-per-el-lipse su-per-el-lip-ses su-per-el-lip-ti-cal " +
  "su-per-el-lip-ti-cally su-per-el-lip-ti-cal-ness sup-ple-mental su-praor-di-nate su-prem-a-cist su-prem-a-cists " +
  "su-prem-acy sur-ger-ies sur-gery sus-que-hanna sus-tain-a-bil-ity sus-tain-a-ble swamina-than swell-ing " +
  "swit-zer-land syl-vester syn-chro-nic-ity syn-o-nym syn-on-y-mous syn-o-nyms syn-on-ymy sys-tem-a-tize " +
  "sys-tem-a-tized szt-acho taban-lar tab-leau tab-leaux taff-rail take-o-ver take-o-vers ta-les tap-es-tries " +
  "tap-es-try ta-rama ta-rekegne ta-rihi tauberian tax-on-o-mies tech-nical tech-nische tech-nol-o-gies tecnolog " +
  "te-davis-inde teh-like tejada tekle-wold tel-e-com tel-e-com-mu-ni-ca-tion tel-e-com-mu-ni-ca-tions " +
  "te-leg-ra-pher te-leg-ra-phers tel-e-ki-netic tel-e-ki-net-ics tel-e-phone tel-ero-bot-ics tel-e-vi-sion " +
  "tell-ers tell-ing te-mas temizleme tem-poral tem-po-rar-ily ten-ellus ten-nes-see ten-ure ter-anodes te-resa " +
  "ter-ri-to-rial test-a-ble testbed test-ers tet-rabu-tylammo-nium textheight textwidth the-a-tre the-o-dore " +
  "ther-mo-e-las-tic ther-mo-e-las-tic-ity ther-mom-e-ter ther-mom-e-ters ther-mo-nu-clear thir-u-vanan-dapu-ram " +
  "thlytras ti-cas ti-cos ti-ers ti-gers timestamp timestamps tirilmesi tirilmi ti-tan-ate tit-ilayo toksik " +
  "tok-sis-ite tol-ches-ter to-man to-maszewski to-mato toolkits top-lam top-lama topo-graph-ical to-pog-ra-phy " +
  "topoi-so-mer-ase topoi-so-mer-ases torres tounic-ode tour-ist toyota trac-ta-ble trans-at-lan-tic trans-ceiv-ers " +
  "trans-cends tran-sept tran-septs trans-for-mation trans-for-mations trans-form-a-tive trans-pa-cific " +
  "trans-par-en-cies trans-par-ency transpile transpiled transpiler transpil-ers transpiles transpil-ing transport " +
  "trans-por-ta-tion tra-pe-zium trap-e-zoid trap-e-zoi-dal trap-e-zoids tra-versal tra-vers-als treach-er-ies " +
  "tre-men-dous tribes-man tri-ethyl-a-mine trifi-dus trig-o-no-met-ric trig-o-nom-e-try trin-i-dad tripa-thi " +
  "tri-plet tri-plets tri-plex tri-plexes tri-po-lar trit-i-cum trontelj tur-key tur-keys turmak-tad turn-a-round " +
  "turn-a-rounds turn-o-ver two-fold ty-pol-o-gies ubbcluj uber-mensch uga-ritic ul-lar uluslararas un-a-ble " +
  "un-am-big-u-ous un-a-void-a-ble un-e-qual un-e-ven un-e-venly un-e-ven-ness un-fa-vor-a-ble uni-formity " +
  "un-im-paired un-in-stan-ti-ated un-ion un-me-di-ated un-pre-dict-a-ble un-pre-dict-a-bly un-published " +
  "un-rea-son-a-ble un-re-li-a-ble un-u-su-ally ur-ban-i-za-tion uru-guay us-ers usi-ta-tis-si-mum uti-lise " +
  "uti-li-za-tion uy-gu-la-ma-lar uy-gu-lanan uz-bek-i-stan va-deden va-guer valores val-u-a-ble var-i-a-bil-ity " +
  "var-i-a-ble var-i-a-bles var-i-ance var-i-ances var-i-ant var-i-ants var-i-a-tion var-i-a-tions var-ies " +
  "va-ri-e-ties var-ughese var-y-ing vau-de-ville ve-dic veg-e-ta-bles ven-e-zuela ven-e-zue-lan ve-nir " +
  "ver-allge-meinerte ver-einigung verov-nik vertei-lungen vi-a-ble vid-eos vid-i-assov vi-etnam viiith viith " +
  "vil-lag-ers vi-na-yan vir-tue vir-tues vi-rus vi-ruses vis-ual vis-u-alizar vis-u-al-i-za-tion " +
  "vis-u-al-i-za-tions vis-u-al-ize vis-u-al-ized vis-u-ally vit-a-mins vi-tecek vi-vek vi-vip-arous vol-a-tile " +
  "vurgula-maktad wad-ding wahr-schein-lich-keit-sthe-o-rie warin-ger warm-est wastewater wave-let wave-lets " +
  "weaponi-sa-tion wea-ponry weap-ons web-log web-logs wegary weighting weight-lifter weight-lift-ing wein-stein " +
  "weld-eki-dan wer-ner wertherian whelan which-ever whisk-ers whitespace whitespaces wik-ipe-dia wild-life wiley " +
  "wil-kinson wil-liam wil-liams wing-dings wing-span wing-spans wis-sen-schaft-lich wolff-ian wolf-ram wood-row " +
  "word-spac-ing work-a-round work-a-rounds worldviews woy-engo wpcon-tent wrap-a-round wrap-a-rounds wretch-edly " +
  "xerox xiang xiiith xviiith xviith xxiiird xxiind yadav yara-lanma-lar yijrk yingyong yol-u-yla yo-ruba " +
  "yo-shi-mura yu-saku zag-majster zah-radnictv zea-land zeman ze-min ze-nodo zentrum zerin-deki zerine zey-inde " +
  "zeylerinde zieg-ler zlemlenmi zo-na-tion zo-ol-o-gischer zo-runlu ";
