#!/usr/bin/env python
#-*- coding: utf-8 -*-

#xylic - 
#This application is free software; you can redistribute
#it and/or modify it under the terms of the GNU General Public License
#defined in the COPYING file

#2010 Charlie Barnes.

import sys
import os
import gtk
import gobject
import mimetypes
import xlrd

class xylicActions():
    def __init__(self):

        self.scores = { 1 : { "score" : 32, "taxa" : ["Abdera affinis"] },
                        2 : { "score" : 8, "taxa" : ["Abdera biflexuosa"] },
                        3 : { "score" : 8, "taxa" : ["Abdera flexuosa"] },
                        4 : { "score" : 16, "taxa" : ["Abdera quadrifasciata"] },
                        5 : { "score" : 16, "taxa" : ["Abdera triguttata"] },
                        6 : { "score" : 4, "taxa" : ["Abraeus globosus"] },
                        7 : { "score" : 8, "taxa" : ["Abraeus granulum"] },
                        8 : { "score" : 2, "taxa" : ["Acalles misellus (= turbatus)", "Acalles turbatus", "Acalles misellus"] },
                        9 : { "score" : 8, "taxa" : ["Acalles roboris"] },
                        10 : { "score" : 8, "taxa" : ["Acanthocinus aedilis"] },
                        11 : { "score" : 24, "taxa" : ["Acritus homoeopathicus"] },
                        12 : { "score" : 2, "taxa" : ["Acrulia inflata"] },
                        13 : { "score" : 32, "taxa" : ["Aderus brevicornis"] },
                        14 : { "score" : 8, "taxa" : ["Aderus oculatus"] },
                        15 : { "score" : 8, "taxa" : ["Aderus populneus"] },
                        16 : { "score" : 16, "taxa" : ["Aeletes atomarius"] },
                        17 : { "score" : 16, "taxa" : ["Agathidium arcticum"] },
                        18 : { "score" : 24, "taxa" : ["Agathidium confusum"] },
                        19 : { "score" : 2, "taxa" : ["Agathidium nigrinum"] },
                        20 : { "score" : 2, "taxa" : ["Agathidium nigripenne"] },
                        21 : { "score" : 16, "taxa" : ["Agathidium pisanum (=badium)", "Agathidium badium", "Agathidium pisanum"] },
                        22 : { "score" : 2, "taxa" : ["Agathidium rotundatum"] },
                        23 : { "score" : 2, "taxa" : ["Agathidium seminulum"] },
                        24 : { "score" : 2, "taxa" : ["Agathidium varians"] },
                        25 : { "score" : 8, "taxa" : ["Agrilus angustulus"] },
                        26 : { "score" : 8, "taxa" : ["Agrilus laticornis"] },
                        27 : { "score" : 8, "taxa" : ["Agrilus pannonicus"] },
                        28 : { "score" : 4, "taxa" : ["Agrilus sinuatus"] },
                        29 : { "score" : 24, "taxa" : ["Agrilus viridis"] },
                        30 : { "score" : 2, "taxa" : ["Alosterna tabacicolor"] },
                        31 : { "score" : 24, "taxa" : ["Amarochara bonnairei"] },
                        32 : { "score" : 2, "taxa" : ["Ampedus balteatus"] },
                        33 : { "score" : 32, "taxa" : ["Ampedus cardinalis"] },
                        34 : { "score" : 16, "taxa" : ["Ampedus cinnabarinus"] },
                        35 : { "score" : 8, "taxa" : ["Ampedus elongatulus"] },
                        36 : { "score" : 32, "taxa" : ["Ampedus nigerrimus"] },
                        37 : { "score" : 8, "taxa" : ["Ampedus nigrinus"] },
                        38 : { "score" : 8, "taxa" : ["Ampedus pomorum"] },
                        39 : { "score" : 8, "taxa" : ["Ampedus quercicola (= pomonae)"] },
                        40 : { "score" : 24, "taxa" : ["Ampedus rufipennis"] },
                        41 : { "score" : 32, "taxa" : ["Ampedus sanguineus"] },
                        42 : { "score" : 16, "taxa" : ["Ampedus sanguinolentus"] },
                        43 : { "score" : 32, "taxa" : ["Ampedus tristis"] },
                        44 : { "score" : 4, "taxa" : ["Anaglyptus mysticus"] },
                        45 : { "score" : 16, "taxa" : ["Anaspis bohemica"] },
                        46 : { "score" : 2, "taxa" : ["Anaspis costai"] },
                        47 : { "score" : 1, "taxa" : ["Anaspis frontalis"] },
                        48 : { "score" : 2, "taxa" : ["Anaspis humeralis"] },
                        49 : { "score" : 2, "taxa" : ["Anaspis lurida"] },
                        50 : { "score" : 16, "taxa" : ["Anaspis melanostoma"] },
                        51 : { "score" : 1, "taxa" : ["Anaspis pulicaria"] },
                        52 : { "score" : 1, "taxa" : ["Anaspis rufilabris"] },
                        53 : { "score" : 24, "taxa" : ["Anaspis septentrionalis (= schilskyana)"] },
                        54 : { "score" : 8, "taxa" : ["Anaspis thoracica"] },
                        55 : { "score" : 24, "taxa" : ["Anastrangalia (= Leptura) sanguinolenta"] },
                        56 : { "score" : 2, "taxa" : ["Anisotoma castanea"] },
                        57 : { "score" : 2, "taxa" : ["Anisotoma glabra"] },
                        58 : { "score" : 2, "taxa" : ["Anisotoma humeralis"] },
                        59 : { "score" : 2, "taxa" : ["Anisotoma orbicularis"] },
                        60 : { "score" : 16, "taxa" : ["Anisoxya fuscula"] },
                        61 : { "score" : 8, "taxa" : ["Anitys rubens"] },
                        62 : { "score" : 8, "taxa" : ["Anobium inexspectatum"] },
                        63 : { "score" : 1, "taxa" : ["Anobium punctatum"] },
                        64 : { "score" : 2, "taxa" : ["Anomognathus cuspidatus"] },
                        65 : { "score" : 24, "taxa" : ["Anoplodera (= Leptura) sexguttata"] },
                        66 : { "score" : 32, "taxa" : ["Anthaxia nitidula"] },
                        67 : { "score" : 4, "taxa" : ["Anthocomus fasciatus"] },
                        68 : { "score" : 8, "taxa" : ["Aplocnemus impressus (=pini)"] },
                        69 : { "score" : 16, "taxa" : ["Aplocnemus nigricornis"] },
                        70 : { "score" : 2, "taxa" : ["Arhopalus rusticus"] },
                        71 : { "score" : 8, "taxa" : ["Aromia moschata"] },
                        72 : { "score" : 2, "taxa" : ["Asemum striatum"] },
                        73 : { "score" : 2, "taxa" : ["Aspidiphorus orbiculatus"] },
                        74 : { "score" : 16, "taxa" : ["Atheta autumnalis"] },
                        75 : { "score" : 16, "taxa" : ["Atheta boletophila"] },
                        76 : { "score" : 16, "taxa" : ["Atheta hansseni"] },
                        77 : { "score" : 2, "taxa" : ["Atheta liturata"] },
                        78 : { "score" : 2, "taxa" : ["Atheta subglabra"] },
                        79 : { "score" : 24, "taxa" : ["Atomaria badia"] },
                        80 : { "score" : 16, "taxa" : ["Atomaria lohsei"] },
                        81 : { "score" : 16, "taxa" : ["Atomaria morio"] },
                        82 : { "score" : 16, "taxa" : ["Atomaria procerula"] },
                        83 : { "score" : 2, "taxa" : ["Atomaria pulchra"] },
                        84 : { "score" : 16, "taxa" : ["Atomaria puncticollis"] },
                        85 : { "score" : 1, "taxa" : ["Atrecus affinis"] },
                        86 : { "score" : 16, "taxa" : ["Aulonium trisulcum"] },
                        87 : { "score" : 24, "taxa" : ["Aulonothroscus brevicollis"] },
                        88 : { "score" : 4, "taxa" : ["Axinotarsus ruficollis"] },
                        89 : { "score" : 32, "taxa" : ["Batrisodes adnexus (=buqueti)"] },
                        90 : { "score" : 32, "taxa" : ["Batrisodes delaporti"] },
                        91 : { "score" : 8, "taxa" : ["Batrisodes venustus"] },
                        92 : { "score" : 2, "taxa" : ["Bibloporus bicolor"] },
                        93 : { "score" : 8, "taxa" : ["Bibloporus minutus"] },
                        94 : { "score" : 4, "taxa" : ["Biphyllus lunatus"] },
                        95 : { "score" : 4, "taxa" : ["Bitoma crenata"] },
                        96 : { "score" : 2, "taxa" : ["Bolitochara lucida"] },
                        97 : { "score" : 8, "taxa" : ["Bolitochara mulsanti"] },
                        98 : { "score" : 8, "taxa" : ["Bolitochara pulchra"] },
                        99 : { "score" : 24, "taxa" : ["Bolitochara reyi"] },
                        100 : { "score" : 16, "taxa" : ["Bolitophagus reticulatus"] },
                        101 : { "score" : 32, "taxa" : ["Bostrichus capucinus"] },
                        102 : { "score" : 32, "taxa" : ["Brachygonus (= Ampedus) ruficeps"] },
                        103 : { "score" : 4, "taxa" : ["Caenoscelis sibirica"] },
                        104 : { "score" : 8, "taxa" : ["Calambus (= Selatosomus) bipustulatus"] },
                        105 : { "score" : 32, "taxa" : ["Carabus intricatus"] },
                        106 : { "score" : 32, "taxa" : ["Cardiophorus gramineus"] },
                        107 : { "score" : 32, "taxa" : ["Cardiophorus ruficollis"] },
                        108 : { "score" : 8, "taxa" : ["Carpophilus sexpustulatus"] },
                        109 : { "score" : 4, "taxa" : ["Cartodere constricta"] },
                        110 : { "score" : 2, "taxa" : ["Caulotrupodes aeneopiceus"] },
                        111 : { "score" : 8, "taxa" : ["Cerylon fagi"] },
                        112 : { "score" : 2, "taxa" : ["Cerylon ferrugineum"] },
                        113 : { "score" : 4, "taxa" : ["Cerylon histeroides"] },
                        114 : { "score" : 16, "taxa" : ["Choragus sheppardi"] },
                        115 : { "score" : 32, "taxa" : ["Chrysanthia nigricornis"] },
                        116 : { "score" : 8, "taxa" : ["Cicones variegata"] },
                        117 : { "score" : 2, "taxa" : ["Cis alni"] },
                        118 : { "score" : 2, "taxa" : ["Cis bidentatus"] },
                        119 : { "score" : 1, "taxa" : ["Cis boleti"] },
                        120 : { "score" : 24, "taxa" : ["Cis coluber"] },
                        121 : { "score" : 24, "taxa" : ["Cis dentatus"] },
                        122 : { "score" : 2, "taxa" : ["Cis fagi"] },
                        123 : { "score" : 2, "taxa" : ["Cis festivus"] },
                        124 : { "score" : 4, "taxa" : ["Cis hispidus"] },
                        125 : { "score" : 8, "taxa" : ["Cis jacquemarti"] },
                        126 : { "score" : 8, "taxa" : ["Cis lineatocribratus"] },
                        127 : { "score" : 4, "taxa" : ["Cis micans"] },
                        128 : { "score" : 2, "taxa" : ["Cis nitidus"] },
                        129 : { "score" : 4, "taxa" : ["Cis punctulatus"] },
                        130 : { "score" : 2, "taxa" : ["Cis pygmaeus"] },
                        131 : { "score" : 2, "taxa" : ["Cis setiger"] },
                        132 : { "score" : 2, "taxa" : ["Cis vestitus"] },
                        133 : { "score" : 1, "taxa" : ["Clytus arietis"] },
                        134 : { "score" : 16, "taxa" : ["Colydium elongatum"] },
                        135 : { "score" : 8, "taxa" : ["Conopalpus testaceus"] },
                        136 : { "score" : 8, "taxa" : ["Corticaria alleni"] },
                        137 : { "score" : 24, "taxa" : ["Corticaria fagi"] },
                        138 : { "score" : 8, "taxa" : ["Corticaria linearis"] },
                        139 : { "score" : 16, "taxa" : ["Corticaria longicollis"] },
                        140 : { "score" : 16, "taxa" : ["Corticaria polypori"] },
                        141 : { "score" : 8, "taxa" : ["Corticeus bicolor"] },
                        142 : { "score" : 24, "taxa" : ["Corticeus unicolor"] },
                        143 : { "score" : 2, "taxa" : ["Coryphium angusticolle"] },
                        144 : { "score" : 16, "taxa" : ["Cossonus linearis"] },
                        145 : { "score" : 8, "taxa" : ["Cossonus parallelepipedus"] },
                        146 : { "score" : 8, "taxa" : ["Cryptarcha strigata"] },
                        147 : { "score" : 8, "taxa" : ["Cryptarcha undata"] },
                        148 : { "score" : 2, "taxa" : ["Cryptolestes duplicatus"] },
                        149 : { "score" : 2, "taxa" : ["Cryptolestes ferrugineus"] },
                        150 : { "score" : 8, "taxa" : ["Cryptophagus acuminatus"] },
                        151 : { "score" : 8, "taxa" : ["Cryptophagus angustus"] },
                        152 : { "score" : 16, "taxa" : ["Cryptophagus confusus"] },
                        153 : { "score" : 24, "taxa" : ["Cryptophagus corticinus"] },
                        154 : { "score" : 1, "taxa" : ["Cryptophagus dentatus"] },
                        155 : { "score" : 24, "taxa" : ["Cryptophagus falcozi"] },
                        156 : { "score" : 16, "taxa" : ["Cryptophagus intermedius"] },
                        157 : { "score" : 8, "taxa" : ["Cryptophagus labilis"] },
                        158 : { "score" : 16, "taxa" : ["Cryptophagus micaceus"] },
                        159 : { "score" : 8, "taxa" : ["Cryptophagus ruficornis"] },
                        160 : { "score" : 4, "taxa" : ["Ctesias serra"] },
                        161 : { "score" : 16, "taxa" : ["Cyanostolus aeneus"] },
                        162 : { "score" : 4, "taxa" : ["Cyphea curtula"] },
                        163 : { "score" : 2, "taxa" : ["Dacne bipustulata"] },
                        164 : { "score" : 2, "taxa" : ["Dacne rufifrons"] },
                        165 : { "score" : 2, "taxa" : ["Dadobia immersa"] },
                        166 : { "score" : 2, "taxa" : ["Dasytes aeratus (= aerosus)"] },
                        167 : { "score" : 16, "taxa" : ["Dasytes niger"] },
                        168 : { "score" : 8, "taxa" : ["Dasytes plumbeus"] },
                        169 : { "score" : 8, "taxa" : ["Dendrophagus crenatus"] },
                        170 : { "score" : 1, "taxa" : ["Denticollis linearis"] },
                        171 : { "score" : 8, "taxa" : ["Dexiogyia corticina"] },
                        172 : { "score" : 24, "taxa" : ["Diaperus boleti"] },
                        173 : { "score" : 16, "taxa" : ["Dictyoptera aurora"] },
                        174 : { "score" : 1, "taxa" : ["Dinaraea aequata"] },
                        175 : { "score" : 2, "taxa" : ["Dinaraea linearis"] },
                        176 : { "score" : 32, "taxa" : ["Dinoptera (= Acmaeops) collaris"] },
                        177 : { "score" : 8, "taxa" : ["Diplocoelus fagi"] },
                        178 : { "score" : 16, "taxa" : ["Dorcatoma ambjourni"] },
                        179 : { "score" : 4, "taxa" : ["Dorcatoma chrysomelina"] },
                        180 : { "score" : 16, "taxa" : ["Dorcatoma dresdensis"] },
                        181 : { "score" : 8, "taxa" : ["Dorcatoma flavicornis"] },
                        182 : { "score" : 16, "taxa" : ["Dorcatoma serra"] },
                        183 : { "score" : 2, "taxa" : ["Dorcus parallelepipedus"] },
                        184 : { "score" : 2, "taxa" : ["Dropephylla devillei (= grandiloqua)"] },
                        185 : { "score" : 8, "taxa" : ["Dropephylla heeri"] },
                        186 : { "score" : 1, "taxa" : ["Dropephylla ioptera"] },
                        187 : { "score" : 1, "taxa" : ["Dropephylla vilis"] },
                        188 : { "score" : 2, "taxa" : ["Dryocoetes autographus"] },
                        189 : { "score" : 16, "taxa" : ["Dryocoetinus alni"] },
                        190 : { "score" : 2, "taxa" : ["Dryocoetinus villosus"] },
                        191 : { "score" : 2, "taxa" : ["Dryophilus pusillus"] },
                        192 : { "score" : 32, "taxa" : ["Dryophthorus corticalis"] },
                        193 : { "score" : 32, "taxa" : ["Elater ferrugineus"] },
                        194 : { "score" : 4, "taxa" : ["Eledona agricola"] },
                        195 : { "score" : 2, "taxa" : ["Emicmus testaceus"] },
                        196 : { "score" : 2, "taxa" : ["Endomychus coccineus"] },
                        197 : { "score" : 32, "taxa" : ["Endophloeus markovichianus"] },
                        198 : { "score" : 8, "taxa" : ["Enicmus brevicornis"] },
                        199 : { "score" : 8, "taxa" : ["Enicmus fungicola"] },
                        200 : { "score" : 8, "taxa" : ["Enicmus rugosus"] },
                        201 : { "score" : 2, "taxa" : ["Ennearthron cornutum"] },
                        202 : { "score" : 16, "taxa" : ["Epierus comptus"] },
                        203 : { "score" : 8, "taxa" : ["Epiphanis cornutus"] },
                        204 : { "score" : 8, "taxa" : ["Epuraea angustula"] },
                        205 : { "score" : 2, "taxa" : ["Epuraea biguttata"] },
                        206 : { "score" : 8, "taxa" : ["Epuraea distincta"] },
                        207 : { "score" : 8, "taxa" : ["Epuraea fuscicollis"] },
                        208 : { "score" : 8, "taxa" : ["Epuraea guttata"] },
                        209 : { "score" : 2, "taxa" : ["Epuraea limbata"] },
                        210 : { "score" : 8, "taxa" : ["Epuraea longula"] },
                        211 : { "score" : 1, "taxa" : ["Epuraea marseuli (= pusilla)"] },
                        212 : { "score" : 24, "taxa" : ["Epuraea neglecta"] },
                        213 : { "score" : 2, "taxa" : ["Epuraea rufomarginata"] },
                        214 : { "score" : 1, "taxa" : ["Epuraea silacea (= deleta)"] },
                        215 : { "score" : 8, "taxa" : ["Epuraea terminalis (= adumbrata)"] },
                        216 : { "score" : 8, "taxa" : ["Epuraea thoracica"] },
                        217 : { "score" : 16, "taxa" : ["Epuraea variegata"] },
                        218 : { "score" : 2, "taxa" : ["Epurea pallescens (= florea)"] },
                        219 : { "score" : 2, "taxa" : ["Ernobius mollis"] },
                        220 : { "score" : 2, "taxa" : ["Ernobius nigrinus"] },
                        221 : { "score" : 16, "taxa" : ["Ernoporus caucasicus"] },
                        222 : { "score" : 8, "taxa" : ["Ernoporus fagi"] },
                        223 : { "score" : 32, "taxa" : ["Ernoporus tiliae"] },
                        224 : { "score" : 32, "taxa" : ["Eucnemis capucina"] },
                        225 : { "score" : 32, "taxa" : ["Euconnus pragensis"] },
                        226 : { "score" : 16, "taxa" : ["Euplectus bescidicus"] },
                        227 : { "score" : 8, "taxa" : ["Euplectus bonvouloiri"] },
                        228 : { "score" : 32, "taxa" : ["Euplectus brunneus"] },
                        229 : { "score" : 8, "taxa" : ["Euplectus fauveli"] },
                        230 : { "score" : 2, "taxa" : ["Euplectus infirmus"] },
                        231 : { "score" : 2, "taxa" : ["Euplectus karsteni"] },
                        232 : { "score" : 8, "taxa" : ["Euplectus kirbyi"] },
                        233 : { "score" : 24, "taxa" : ["Euplectus nanus"] },
                        234 : { "score" : 2, "taxa" : ["Euplectus piceus"] },
                        235 : { "score" : 24, "taxa" : ["Euplectus punctatus"] },
                        236 : { "score" : 24, "taxa" : ["Euryusa optabilis"] },
                        237 : { "score" : 24, "taxa" : ["Euryusa sinuata"] },
                        238 : { "score" : 32, "taxa" : ["Eutheia formicetorum"] },
                        239 : { "score" : 32, "taxa" : ["Eutheia linearis"] },
                        240 : { "score" : 1, "taxa" : ["Gabrius splendidulus"] },
                        241 : { "score" : 32, "taxa" : ["Gastrallus immarginatus"] },
                        242 : { "score" : 2, "taxa" : ["Glischrochilus quadriguttatus"] },
                        243 : { "score" : 2, "taxa" : ["Glischrochilus quadripunctatus"] },
                        244 : { "score" : 32, "taxa" : ["Globicornis rufitarsis (=nigripes)"] },
                        245 : { "score" : 32, "taxa" : ["Gnorimus nobilis"] },
                        246 : { "score" : 32, "taxa" : ["Gnorimus variabilis"] },
                        247 : { "score" : 2, "taxa" : ["Gonodera luperus"] },
                        248 : { "score" : 1, "taxa" : ["Grammoptera ruficornis"] },
                        249 : { "score" : 24, "taxa" : ["Grammoptera ustulata"] },
                        250 : { "score" : 24, "taxa" : ["Grammoptera variegata"] },
                        251 : { "score" : 2, "taxa" : ["Grynobius planus"] },
                        252 : { "score" : 8, "taxa" : ["Gyrophaena angustata"] },
                        253 : { "score" : 2, "taxa" : ["Gyrophaena bihamata"] },
                        254 : { "score" : 8, "taxa" : ["Gyrophaena congrua"] },
                        255 : { "score" : 8, "taxa" : ["Gyrophaena joyi"] },
                        256 : { "score" : 2, "taxa" : ["Gyrophaena latissima"] },
                        257 : { "score" : 8, "taxa" : ["Gyrophaena lucidula"] },
                        258 : { "score" : 2, "taxa" : ["Gyrophaena minima"] },
                        259 : { "score" : 16, "taxa" : ["Gyrophaena munsteri"] },
                        260 : { "score" : 16, "taxa" : ["Gyrophaena poweri"] },
                        261 : { "score" : 24, "taxa" : ["Gyrophaena pseudonana"] },
                        262 : { "score" : 16, "taxa" : ["Gyrophaena pulchella"] },
                        263 : { "score" : 8, "taxa" : ["Gyrophaena strictula"] },
                        264 : { "score" : 8, "taxa" : ["Hadrobregmus denticollis"] },
                        265 : { "score" : 8, "taxa" : ["Hallomenus binotatus"] },
                        266 : { "score" : 2, "taxa" : ["Hapalaraea pygmaea"] },
                        267 : { "score" : 2, "taxa" : ["Haploglossa gentilis"] },
                        268 : { "score" : 8, "taxa" : ["Haploglossa marginalis"] },
                        269 : { "score" : 8, "taxa" : ["Harminius undulatus"] },
                        270 : { "score" : 8, "taxa" : ["Helops caeruleus"] },
                        271 : { "score" : 1, "taxa" : ["Hemicoelus fulvicornis"] },
                        272 : { "score" : 24, "taxa" : ["Hemicoelus nitidus"] },
                        273 : { "score" : 2, "taxa" : ["Henoticus serratus"] },
                        274 : { "score" : 2, "taxa" : ["Homalota plana"] },
                        275 : { "score" : 1, "taxa" : ["Hylastes ater"] },
                        276 : { "score" : 2, "taxa" : ["Hylastes brunneus"] },
                        277 : { "score" : 2, "taxa" : ["Hylastes opacus"] },
                        278 : { "score" : 4, "taxa" : ["Hylecoetus dermestoides"] },
                        279 : { "score" : 1, "taxa" : ["Hylesinus (= Leperisinus) varius"] },
                        280 : { "score" : 2, "taxa" : ["Hylesinus crenatus"] },
                        281 : { "score" : 2, "taxa" : ["Hylesinus oleiperda"] },
                        282 : { "score" : 8, "taxa" : ["Hylesinus orni"] },
                        283 : { "score" : 32, "taxa" : ["Hylis cariniceps"] },
                        284 : { "score" : 24, "taxa" : ["Hylis olexai"] },
                        285 : { "score" : 1, "taxa" : ["Hylobius abietis"] },
                        286 : { "score" : 1, "taxa" : ["Hylurgops palliatus"] },
                        287 : { "score" : 32, "taxa" : ["Hypebaeus flavipes"] },
                        288 : { "score" : 16, "taxa" : ["Hypulus quercinus"] },
                        289 : { "score" : 2, "taxa" : ["Ips acuminatus"] },
                        290 : { "score" : 16, "taxa" : ["Ischnodes sanguinicollis"] },
                        291 : { "score" : 16, "taxa" : ["Ischnoglossa obscura"] },
                        292 : { "score" : 2, "taxa" : ["Ischnoglossa prolixa"] },
                        293 : { "score" : 2, "taxa" : ["Ischnoglossa turcica"] },
                        294 : { "score" : 24, "taxa" : ["Ischnomera caerulea"] },
                        295 : { "score" : 32, "taxa" : ["Ischnomera cinerascens"] },
                        296 : { "score" : 4, "taxa" : ["Ischnomera cyanea"] },
                        297 : { "score" : 8, "taxa" : ["Ischnomera sanguinicollis"] },
                        298 : { "score" : 24, "taxa" : ["Judolia sexmaculata"] },
                        299 : { "score" : 8, "taxa" : ["Kissophagus hederae"] },
                        300 : { "score" : 8, "taxa" : ["Korynetes caeruleus"] },
                        301 : { "score" : 32, "taxa" : ["Lacon quercus"] },
                        302 : { "score" : 32, "taxa" : ["Laemophloeus monilis"] },
                        303 : { "score" : 32, "taxa" : ["Lamia textor"] },
                        304 : { "score" : 8, "taxa" : ["Lathridius consimilis"] },
                        305 : { "score" : 2, "taxa" : ["Leiopus nebulosus"] },
                        306 : { "score" : 16, "taxa" : ["Leptura (= Strangalia) aurulenta"] },
                        307 : { "score" : 2, "taxa" : ["Leptura (= Strangalia) quadrifasciata"] },
                        308 : { "score" : 32, "taxa" : ["Lepturobosca virens"] },
                        309 : { "score" : 1, "taxa" : ["Leptusa fumida"] },
                        310 : { "score" : 8, "taxa" : ["Leptusa norvegica"] },
                        311 : { "score" : 2, "taxa" : ["Leptusa pulchella"] },
                        312 : { "score" : 1, "taxa" : ["Leptusa ruficollis"] },
                        313 : { "score" : 32, "taxa" : ["Limoniscus violaceus"] },
                        314 : { "score" : 16, "taxa" : ["Lissodema cursor"] },
                        315 : { "score" : 8, "taxa" : ["Lissodema quadripustulata"] },
                        316 : { "score" : 2, "taxa" : ["Litargus connexus"] },
                        317 : { "score" : 8, "taxa" : ["Lucanus cervus"] },
                        318 : { "score" : 4, "taxa" : ["Lyctus brunneus"] },
                        319 : { "score" : 8, "taxa" : ["Lyctus linearis"] },
                        320 : { "score" : 32, "taxa" : ["Lymantor coryli"] },
                        321 : { "score" : 32, "taxa" : ["Lymexylon navale"] },
                        322 : { "score" : 2, "taxa" : ["Magdalis armigera"] },
                        323 : { "score" : 8, "taxa" : ["Magdalis barbicornis"] },
                        324 : { "score" : 4, "taxa" : ["Magdalis carbonaria"] },
                        325 : { "score" : 4, "taxa" : ["Magdalis cerasi"] },
                        326 : { "score" : 16, "taxa" : ["Magdalis duplicata"] },
                        327 : { "score" : 8, "taxa" : ["Magdalis phlegmatica"] },
                        328 : { "score" : 2, "taxa" : ["Magdalis ruficornis"] },
                        329 : { "score" : 1, "taxa" : ["Malachius bipustulatus"] },
                        330 : { "score" : 8, "taxa" : ["Malthinus balteatus"] },
                        331 : { "score" : 1, "taxa" : ["Malthinus flaveolus"] },
                        332 : { "score" : 8, "taxa" : ["Malthinus frontalis"] },
                        333 : { "score" : 2, "taxa" : ["Malthinus seriepunctatus"] },
                        334 : { "score" : 24, "taxa" : ["Malthodes crassicornis"] },
                        335 : { "score" : 2, "taxa" : ["Malthodes dispar"] },
                        336 : { "score" : 8, "taxa" : ["Malthodes fibulatus"] },
                        337 : { "score" : 2, "taxa" : ["Malthodes flavoguttatus"] },
                        338 : { "score" : 2, "taxa" : ["Malthodes fuscus"] },
                        339 : { "score" : 8, "taxa" : ["Malthodes guttifer"] },
                        340 : { "score" : 1, "taxa" : ["Malthodes marginatus"] },
                        341 : { "score" : 16, "taxa" : ["Malthodes maurus"] },
                        342 : { "score" : 1, "taxa" : ["Malthodes minimus"] },
                        343 : { "score" : 2, "taxa" : ["Malthodes mysticus"] },
                        344 : { "score" : 2, "taxa" : ["Malthodes pumilus"] },
                        345 : { "score" : 32, "taxa" : ["Megapenthes lugens"] },
                        346 : { "score" : 16, "taxa" : ["Megarthrus hemipterus"] },
                        347 : { "score" : 8, "taxa" : ["Megatoma undata"] },
                        348 : { "score" : 32, "taxa" : ["Melandrya barbata"] },
                        349 : { "score" : 4, "taxa" : ["Melandrya caraboides"] },
                        350 : { "score" : 2, "taxa" : ["Melanophila acuminata"] },
                        351 : { "score" : 1, "taxa" : ["Melanotus villosus (= erythropus)"] },
                        352 : { "score" : 4, "taxa" : ["Melasis buprestoides"] },
                        353 : { "score" : 8, "taxa" : ["Mesites tardii"] },
                        354 : { "score" : 24, "taxa" : ["Mesosa nebulosa"] },
                        355 : { "score" : 16, "taxa" : ["Micrambe bimaculatus"] },
                        356 : { "score" : 16, "taxa" : ["Micridium halidaii"] },
                        357 : { "score" : 8, "taxa" : ["Microrhagus (= Dirhagus) pygmaeus"] },
                        358 : { "score" : 24, "taxa" : ["Microscydmus minimus"] },
                        359 : { "score" : 16, "taxa" : ["Molorchus umbellatarum"] },
                        360 : { "score" : 16, "taxa" : ["Mordellistena neuwaldeggiana"] },
                        361 : { "score" : 8, "taxa" : ["Mordellistena variegata"] },
                        362 : { "score" : 4, "taxa" : ["Mordellochroa abdominalis"] },
                        363 : { "score" : 2, "taxa" : ["Mycetaea subterranea (= hirta)"] },
                        364 : { "score" : 16, "taxa" : ["Mycetochara humeralis"] },
                        365 : { "score" : 2, "taxa" : ["Mycetophagus atomarius"] },
                        366 : { "score" : 32, "taxa" : ["Mycetophagus fulvicollis"] },
                        367 : { "score" : 2, "taxa" : ["Mycetophagus multipunctatus"] },
                        368 : { "score" : 4, "taxa" : ["Mycetophagus piceus"] },
                        369 : { "score" : 16, "taxa" : ["Mycetophagus populi"] },
                        370 : { "score" : 16, "taxa" : ["Mycetophagus quadriguttatus"] },
                        371 : { "score" : 2, "taxa" : ["Mycetophagus quadripustulatus"] },
                        372 : { "score" : 2, "taxa" : ["Nemadus colonoides"] },
                        373 : { "score" : 24, "taxa" : ["Nemozoma elongatum"] },
                        374 : { "score" : 8, "taxa" : ["Neuraphes plicicollis"] },
                        375 : { "score" : 8, "taxa" : ["Nossidium pilosellum"] },
                        376 : { "score" : 16, "taxa" : ["Notolaemus unifasciatus"] },
                        377 : { "score" : 2, "taxa" : ["Nudobius lentus"] },
                        378 : { "score" : 32, "taxa" : ["Obrium cantharinum"] },
                        379 : { "score" : 2, "taxa" : ["Ochina ptinoides"] },
                        380 : { "score" : 1, "taxa" : ["Octotemnus glabriculus"] },
                        381 : { "score" : 24, "taxa" : ["Oedemera virescens"] },
                        382 : { "score" : 8, "taxa" : ["Oncomera femorata"] },
                        383 : { "score" : 8, "taxa" : ["Opilio mollis"] },
                        384 : { "score" : 4, "taxa" : ["Orchesia micans"] },
                        385 : { "score" : 8, "taxa" : ["Orchesia minor"] },
                        386 : { "score" : 4, "taxa" : ["Orchesia undulata"] },
                        387 : { "score" : 16, "taxa" : ["Orthoperus aequalis (= nitidulus)"] },
                        388 : { "score" : 4, "taxa" : ["Orthoperus mundus"] },
                        389 : { "score" : 2, "taxa" : ["Orthotomicus suturalis"] },
                        390 : { "score" : 16, "taxa" : ["Osphya bipunctata"] },
                        391 : { "score" : 32, "taxa" : ["Ostoma ferrugineum"] },
                        392 : { "score" : 32, "taxa" : ["Oxylaemus cylindricus"] },
                        393 : { "score" : 24, "taxa" : ["Oxylaemus variolosus"] },
                        394 : { "score" : 2, "taxa" : ["Pachytodes (= Judolia) cerambyciformis"] },
                        395 : { "score" : 24, "taxa" : ["Paracorymbia (= Leptura) fulva"] },
                        396 : { "score" : 16, "taxa" : ["Paranopleta inhabilis"] },
                        397 : { "score" : 2, "taxa" : ["Paromalus flavicornis"] },
                        398 : { "score" : 32, "taxa" : ["Paromalus parallelepipedus"] },
                        399 : { "score" : 16, "taxa" : ["Pediacus depressus"] },
                        400 : { "score" : 4, "taxa" : ["Pediacus dermestoides"] },
                        401 : { "score" : 32, "taxa" : ["Pedostrangalia (=Leptura) revestita"] },
                        402 : { "score" : 4, "taxa" : ["Pentaphyllus testaceus"] },
                        403 : { "score" : 2, "taxa" : ["Philonthus subuliformis"] },
                        404 : { "score" : 2, "taxa" : ["Phloeocharis subtillissima"] },
                        405 : { "score" : 24, "taxa" : ["Phloeodroma concolor"] },
                        406 : { "score" : 2, "taxa" : ["Phloeonomus punctipennis"] },
                        407 : { "score" : 2, "taxa" : ["Phloeonomus pusillus"] },
                        408 : { "score" : 32, "taxa" : ["Phloeophagus (= Rhyncholus) gracilis"] },
                        409 : { "score" : 2, "taxa" : ["Phloeophagus (= Rhyncholus) lignarius"] },
                        410 : { "score" : 2, "taxa" : ["Phloeopora bernhaueri (= teres)"] },
                        411 : { "score" : 8, "taxa" : ["Phloeopora corticalis (= angustiformis)"] },
                        412 : { "score" : 1, "taxa" : ["Phloeopora testacea"] },
                        413 : { "score" : 2, "taxa" : ["Phloeostiba lapponica"] },
                        414 : { "score" : 2, "taxa" : ["Phloeostiba plana"] },
                        415 : { "score" : 8, "taxa" : ["Phloiophilus edwardsi"] },
                        416 : { "score" : 8, "taxa" : ["Phloiotrya vaudoueri"] },
                        417 : { "score" : 24, "taxa" : ["Phyllodrepa nigra"] },
                        418 : { "score" : 8, "taxa" : ["Phyllodrepoidea crenata"] },
                        419 : { "score" : 16, "taxa" : ["Phymatodes alni"] },
                        420 : { "score" : 4, "taxa" : ["Phymatodes testaceus"] },
                        421 : { "score" : 2, "taxa" : ["Pissodes castaneus"] },
                        422 : { "score" : 2, "taxa" : ["Pissodes pini"] },
                        423 : { "score" : 2, "taxa" : ["Pityogenes bidentatus"] },
                        424 : { "score" : 16, "taxa" : ["Pityogenes quadridens"] },
                        425 : { "score" : 8, "taxa" : ["Pityogenes trepanatus"] },
                        426 : { "score" : 2, "taxa" : ["Pityophagus ferrugineus"] },
                        427 : { "score" : 24, "taxa" : ["Pityophthorus lichtensteini"] },
                        428 : { "score" : 2, "taxa" : ["Pityophthorus pubescens"] },
                        429 : { "score" : 8, "taxa" : ["Placusa depressa"] },
                        430 : { "score" : 2, "taxa" : ["Placusa pumilio"] },
                        431 : { "score" : 8, "taxa" : ["Placusa tachyporoides"] },
                        432 : { "score" : 32, "taxa" : ["Plagionatus arcuatus"] },
                        433 : { "score" : 24, "taxa" : ["Platycis cosnardi"] },
                        434 : { "score" : 8, "taxa" : ["Platycis minutus"] },
                        435 : { "score" : 32, "taxa" : ["Platydema violaceum"] },
                        436 : { "score" : 8, "taxa" : ["Platypus cylindrus"] },
                        437 : { "score" : 4, "taxa" : ["Platyrhinus resinosus"] },
                        438 : { "score" : 8, "taxa" : ["Platystomos albinus"] },
                        439 : { "score" : 32, "taxa" : ["Plectophloeus nitidus"] },
                        440 : { "score" : 8, "taxa" : ["Plegaderus dissectus"] },
                        441 : { "score" : 16, "taxa" : ["Pogonocherus fasciculatus"] },
                        442 : { "score" : 2, "taxa" : ["Pogonocherus hispidulus"] },
                        443 : { "score" : 2, "taxa" : ["Pogonocherus hispidus"] },
                        444 : { "score" : 8, "taxa" : ["Prionocyphon serricornis"] },
                        445 : { "score" : 16, "taxa" : ["Prionus coriarius"] },
                        446 : { "score" : 8, "taxa" : ["Prionychus ater"] },
                        447 : { "score" : 32, "taxa" : ["Prionychus melanarius"] },
                        448 : { "score" : 16, "taxa" : ["Procraerus tibialis"] },
                        449 : { "score" : 8, "taxa" : ["Pseudocistela ceramboides"] },
                        450 : { "score" : 4, "taxa" : ["Pseudotriphyllus suturalis"] },
                        451 : { "score" : 2, "taxa" : ["Pteleobius vittatus"] },
                        452 : { "score" : 8, "taxa" : ["Ptenidium gressneri"] },
                        453 : { "score" : 16, "taxa" : ["Ptenidium turgidum"] },
                        454 : { "score" : 2, "taxa" : ["Pteryx suturalis"] },
                        455 : { "score" : 1, "taxa" : ["Ptilinus pectinicornis"] },
                        456 : { "score" : 16, "taxa" : ["Ptiliolum caledonicum"] },
                        457 : { "score" : 2, "taxa" : ["Ptinella aptera"] },
                        458 : { "score" : 8, "taxa" : ["Ptinella denticollis"] },
                        459 : { "score" : 16, "taxa" : ["Ptinella limbata"] },
                        460 : { "score" : 8, "taxa" : ["Ptinomorphus (= Hedobia) imperialis"] },
                        461 : { "score" : 24, "taxa" : ["Ptinus lichenum"] },
                        462 : { "score" : 16, "taxa" : ["Ptinus palliatus"] },
                        463 : { "score" : 8, "taxa" : ["Ptinus subpilosus"] },
                        464 : { "score" : 4, "taxa" : ["Pyrochroa coccinea"] },
                        465 : { "score" : 1, "taxa" : ["Pyrochroa serraticornis"] },
                        466 : { "score" : 16, "taxa" : ["Pyropterus nigroruber"] },
                        467 : { "score" : 24, "taxa" : ["Pyrrhidium sanguineum"] },
                        468 : { "score" : 16, "taxa" : ["Pytho depressus"] },
                        469 : { "score" : 16, "taxa" : ["Quedius aetolicus"] },
                        470 : { "score" : 8, "taxa" : ["Quedius brevicornis"] },
                        471 : { "score" : 4, "taxa" : ["Quedius maurus"] },
                        472 : { "score" : 8, "taxa" : ["Quedius microps"] },
                        473 : { "score" : 2, "taxa" : ["Quedius plagiatus"] },
                        474 : { "score" : 8, "taxa" : ["Quedius scitus"] },
                        475 : { "score" : 8, "taxa" : ["Quedius truncicola (=ventralis)"] },
                        476 : { "score" : 4, "taxa" : ["Quedius xanthopus"] },
                        477 : { "score" : 16, "taxa" : ["Rabocerus foveolatus"] },
                        478 : { "score" : 8, "taxa" : ["Rabocerus gabrieli"] },
                        479 : { "score" : 1, "taxa" : ["Rhagium bifasciatum"] },
                        480 : { "score" : 8, "taxa" : ["Rhagium inquisitor"] },
                        481 : { "score" : 1, "taxa" : ["Rhagium mordax"] },
                        482 : { "score" : 1, "taxa" : ["Rhinosimus planirostris"] },
                        483 : { "score" : 1, "taxa" : ["Rhinosimus ruficollis"] },
                        484 : { "score" : 1, "taxa" : ["Rhizophagus bipustulatus"] },
                        485 : { "score" : 2, "taxa" : ["Rhizophagus cribratus"] },
                        486 : { "score" : 2, "taxa" : ["Rhizophagus depressus"] },
                        487 : { "score" : 1, "taxa" : ["Rhizophagus dispar"] },
                        488 : { "score" : 2, "taxa" : ["Rhizophagus ferrugineus"] },
                        489 : { "score" : 4, "taxa" : ["Rhizophagus nitidulus"] },
                        490 : { "score" : 24, "taxa" : ["Rhizophagus oblongicollis"] },
                        491 : { "score" : 2, "taxa" : ["Rhizophagus parallelocollis"] },
                        492 : { "score" : 24, "taxa" : ["Rhizophagus parvulus"] },
                        493 : { "score" : 2, "taxa" : ["Rhizophagus perforatus"] },
                        494 : { "score" : 16, "taxa" : ["Rhizophagus picipes"] },
                        495 : { "score" : 24, "taxa" : ["Rhopalodontus perforatus"] },
                        496 : { "score" : 8, "taxa" : ["Rhyncolus chloropus (=Eremotes ater)"] },
                        497 : { "score" : 1, "taxa" : ["Rutpela (= Strangalia) maculata"] },
                        498 : { "score" : 2, "taxa" : ["Salpingus ater"] },
                        499 : { "score" : 2, "taxa" : ["Salpingus castaneus"] },
                        500 : { "score" : 2, "taxa" : ["Salpingus reyi"] },
                        501 : { "score" : 16, "taxa" : ["Saperda carcharias"] },
                        502 : { "score" : 8, "taxa" : ["Saperda scalaris"] },
                        503 : { "score" : 2, "taxa" : ["Scaphidium quadrimaculatum"] },
                        504 : { "score" : 2, "taxa" : ["Scaphisoma agaricinum"] },
                        505 : { "score" : 24, "taxa" : ["Scaphisoma assimile"] },
                        506 : { "score" : 8, "taxa" : ["Scaphisoma boleti"] },
                        507 : { "score" : 16, "taxa" : ["Schizotus pectinicornis"] },
                        508 : { "score" : 2, "taxa" : ["Scolytus intricatus"] },
                        509 : { "score" : 8, "taxa" : ["Scolytus mali"] },
                        510 : { "score" : 1, "taxa" : ["Scolytus multistriatus"] },
                        511 : { "score" : 8, "taxa" : ["Scolytus ratzeburgi"] },
                        512 : { "score" : 2, "taxa" : ["Scolytus rugulosus"] },
                        513 : { "score" : 2, "taxa" : ["Scolytus scolytus"] },
                        514 : { "score" : 32, "taxa" : ["Scraptia dubia"] },
                        515 : { "score" : 32, "taxa" : ["Scraptia fuscula"] },
                        516 : { "score" : 16, "taxa" : ["Scraptia testacea"] },
                        517 : { "score" : 24, "taxa" : ["Scydmaenus rufus"] },
                        518 : { "score" : 8, "taxa" : ["Sepedophilus bipunctatus"] },
                        519 : { "score" : 2, "taxa" : ["Sepedophilus littoreus"] },
                        520 : { "score" : 2, "taxa" : ["Sepedophilus lusitanicus"] },
                        521 : { "score" : 8, "taxa" : ["Sepedophilus testaceus"] },
                        522 : { "score" : 2, "taxa" : ["Siagonum quadricorne"] },
                        523 : { "score" : 8, "taxa" : ["Silusa rubiginosa"] },
                        524 : { "score" : 32, "taxa" : ["Silvanoprus fagi"] },
                        525 : { "score" : 8, "taxa" : ["Silvanus bidentatus"] },
                        526 : { "score" : 4, "taxa" : ["Silvanus unidentatus"] },
                        527 : { "score" : 2, "taxa" : ["Sinodendron cylindricum"] },
                        528 : { "score" : 2, "taxa" : ["Soronia grisea"] },
                        529 : { "score" : 2, "taxa" : ["Soronia punctatissima"] },
                        530 : { "score" : 8, "taxa" : ["Sphindus dubius"] },
                        531 : { "score" : 16, "taxa" : ["Sphinginus lobatus"] },
                        532 : { "score" : 4, "taxa" : ["Stenagostus rhombeus (= villosus)"] },
                        533 : { "score" : 4, "taxa" : ["Stenichnus bicolor"] },
                        534 : { "score" : 24, "taxa" : ["Stenichnus godarti"] },
                        535 : { "score" : 2, "taxa" : ["Stenocorus meridianus"] },
                        536 : { "score" : 8, "taxa" : ["Stenostola dubia"] },
                        537 : { "score" : 2, "taxa" : ["Stenurella (= Strangalia) melanura"] },
                        538 : { "score" : 24, "taxa" : ["Stenurella (= Strangalia) nigra"] },
                        539 : { "score" : 4, "taxa" : ["Stephostethus alternans"] },
                        540 : { "score" : 16, "taxa" : ["Stereocorynes (= Rhyncholus) truncorum"] },
                        541 : { "score" : 24, "taxa" : ["Stichoglossa semirufa"] },
                        542 : { "score" : 16, "taxa" : ["Stictoleptura (=Anoplodera) scutellata"] },
                        543 : { "score" : 2, "taxa" : ["Sulcacis affinis"] },
                        544 : { "score" : 8, "taxa" : ["Sulcacis bicornis"] },
                        545 : { "score" : 8, "taxa" : ["Symbiotes latus"] },
                        546 : { "score" : 8, "taxa" : ["Synchita humeralis"] },
                        547 : { "score" : 24, "taxa" : ["Synchita separanda"] },
                        548 : { "score" : 32, "taxa" : ["Tachinus bipustulatus"] },
                        549 : { "score" : 32, "taxa" : ["Tachyusida gracilis"] },
                        550 : { "score" : 8, "taxa" : ["Taphrorhychus bicolor"] },
                        551 : { "score" : 32, "taxa" : ["Tarsostenus univittatus"] },
                        552 : { "score" : 32, "taxa" : ["Teredus cylindricus"] },
                        553 : { "score" : 32, "taxa" : ["Teretrius fabricii"] },
                        554 : { "score" : 8, "taxa" : ["Tetratoma ancora"] },
                        555 : { "score" : 16, "taxa" : ["Tetratoma desmaresti"] },
                        556 : { "score" : 2, "taxa" : ["Tetratoma fungorum"] },
                        557 : { "score" : 2, "taxa" : ["Tetrops praeusta"] },
                        558 : { "score" : 16, "taxa" : ["Tetrops starkii"] },
                        559 : { "score" : 2, "taxa" : ["Thamiaraea cinnamomea"] },
                        560 : { "score" : 8, "taxa" : ["Thamiaraea hospita"] },
                        561 : { "score" : 4, "taxa" : ["Thanasimus formicarius"] },
                        562 : { "score" : 24, "taxa" : ["Thanasimus rufipes"] },
                        563 : { "score" : 8, "taxa" : ["Thymalus limbatus"] },
                        564 : { "score" : 32, "taxa" : ["Tilloidea unifasciatus"] },
                        565 : { "score" : 8, "taxa" : ["Tillus elongatus"] },
                        566 : { "score" : 24, "taxa" : ["Tomicus minor"] },
                        567 : { "score" : 1, "taxa" : ["Tomicus piniperda"] },
                        568 : { "score" : 16, "taxa" : ["Tomoxia bucephala (= biguttata)"] },
                        569 : { "score" : 8, "taxa" : ["Trachodes hispidus"] },
                        570 : { "score" : 2, "taxa" : ["Trichius fasciatus"] },
                        571 : { "score" : 32, "taxa" : ["Trichonyx sulcicollis"] },
                        572 : { "score" : 24, "taxa" : ["Trinodes hirtus"] },
                        573 : { "score" : 4, "taxa" : ["Triphyllus bicolor"] },
                        574 : { "score" : 2, "taxa" : ["Triplax aenea"] },
                        575 : { "score" : 24, "taxa" : ["Triplax lacordairii"] },
                        576 : { "score" : 4, "taxa" : ["Triplax russica"] },
                        577 : { "score" : 32, "taxa" : ["Triplax scutellaris"] },
                        578 : { "score" : 16, "taxa" : ["Tritoma bipustulata"] },
                        579 : { "score" : 32, "taxa" : ["Tropideres niveirostris"] },
                        580 : { "score" : 32, "taxa" : ["Tropideres sepicola"] },
                        581 : { "score" : 2, "taxa" : ["Trypodendron (= Xyloterus) domesticum"] },
                        582 : { "score" : 2, "taxa" : ["Trypodendron (= Xyloterus) lineatum"] },
                        583 : { "score" : 8, "taxa" : ["Trypodendron (= Xyloterus) signatum"] },
                        584 : { "score" : 16, "taxa" : ["Trypophloeus binodulus (= asperatus)"] },
                        585 : { "score" : 32, "taxa" : ["Trypophloeus granulatus"] },
                        586 : { "score" : 16, "taxa" : ["Uleiota planata"] },
                        587 : { "score" : 32, "taxa" : ["Velleius dilatatus"] },
                        588 : { "score" : 2, "taxa" : ["Vicenzellus ruficollis"] },
                        589 : { "score" : 16, "taxa" : ["Xantholinus angularis"] },
                        590 : { "score" : 4, "taxa" : ["Xestobium rufovillosum"] },
                        591 : { "score" : 4, "taxa" : ["Xyleborinus saxeseni"] },
                        592 : { "score" : 8, "taxa" : ["Xyleborus dispar"] },
                        593 : { "score" : 8, "taxa" : ["Xyleborus dryographus"] },
                        594 : { "score" : 32, "taxa" : ["Xyletinus longitarsus"] },
                        595 : { "score" : 16, "taxa" : ["Xylita laevigata"] },
                        596 : { "score" : 32, "taxa" : ["Xylodromus testaceus"] },
                        597 : { "score" : 8, "taxa" : ["Xylostiba monilicornis"] },
                        598 : { "score" : 8, "taxa" : ["Zilora ferruginea"] },
                       }
                        
        #Load the widget tree
        builder = ""
        self.builder = gtk.Builder()
        self.builder.add_from_string(builder, len(builder))
        self.builder.add_from_file("ui.xml")

        signals = {
                   "mainQuit":self.main_quit,
                   "showAboutDialog":self.show_about_dialog,
                   "parse":self.parse,
                   "selectFile":self.select_file,
                  }
        self.builder.connect_signals(signals)

        treeview = self.builder.get_object("treeview1")
        model = gtk.ListStore(str, int, int, int, str)
        treeview.set_headers_visible(True)
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Site", cell, text=0)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(0)
        treeview.append_column(column)
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Species", cell, text=1)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(1)
        treeview.append_column(column)    
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Scoring Species", cell, text=2)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(2)
        treeview.append_column(column)    
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("SQS", cell, text=3)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(3)
        treeview.append_column(column)    
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("SQI", cell, text=4)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(4)
        treeview.append_column(column)    
            
        treeview.set_model(model)

        #Setup the main window
        self.main_window = self.builder.get_object("window1")
        self.main_window.show()
              
    def select_file(self, widget):
        filetype = mimetypes.guess_type(self.builder.get_object("filechooserbutton2").get_filename())[0]
        
        if filetype == "application/vnd.ms-excel":
            self.parse(widget)
              
    def parse(self, widget):

        cursor = gtk.gdk.Cursor(gtk.gdk.WATCH)
        self.builder.get_object("window1").window.set_cursor(cursor)
    
        while gtk.events_pending():
            gtk.main_iteration()
                    
        treeview = self.builder.get_object("treeview1")
        model = treeview.get_model()
        model.clear()
        
        filename = self.builder.get_object("filechooserbutton2").get_filename()
        filetype = mimetypes.guess_type(filename)[0]
        
        if filetype == "application/vnd.ms-excel":
            book = xlrd.open_workbook(filename)
            
            if book.nsheets > 1:
                
                dialog = self.builder.get_object("dialog1")

                try:
                    self.builder.get_object("hbox5").get_children()[1].destroy()           
                except IndexError:
                    pass
                    
                combobox = gtk.combo_box_new_text()
                
                for name in book.sheet_names():
                    combobox.append_text(name)
                    
                combobox.set_active(0)
                combobox.show()
                self.builder.get_object("hbox5").add(combobox)
                
                self.builder.get_object("window1").window.set_cursor(None)
            
                while gtk.events_pending():
                    gtk.main_iteration()
                
                response = dialog.run()

                if response == 1:
                    sheet = book.sheet_by_name(combobox.get_active_text())
                else:
                    dialog.hide()
                    return -1
                    
                dialog.hide()
                
            else:
                sheet = book.sheet_by_index(0)

            self.builder.get_object("vbox1").set_sensitive(False)
            
            cursor = gtk.gdk.Cursor(gtk.gdk.WATCH)
            self.builder.get_object("window1").window.set_cursor(cursor)
        
            while gtk.events_pending():
                gtk.main_iteration()
                
            for col_index in range(sheet.ncols):
                if sheet.cell(0, col_index).value == "Site":
                    site_position = col_index
                elif sheet.cell(0, col_index).value.lower() == "location":
                    site_position = col_index
                elif sheet.cell(0, col_index).value == "Species":
                    taxon_position = col_index
                elif sheet.cell(0, col_index).value == "Taxon":
                    taxon_position = col_index
                elif sheet.cell(0, col_index).value == "Taxon Name":
                    taxon_position = col_index
                elif sheet.cell(0, col_index).value == "Date":
                    date_position = col_index

            data = {}
            
            for row_index in range(1, sheet.nrows):
                site = sheet.cell(row_index, site_position).value
                taxon = sheet.cell(row_index, taxon_position).value
                    
                if data.has_key(site) and taxon not in data[site]["species_list"]:
                    data[site]["species_list"].append(taxon)
                elif not data.has_key(site):
                    data[site] = { }
                    data[site]["species_list"] = [taxon, ]
                    data[site]["scoring_species"] = [ ]
                    data[site]["sqs"] = 0
                    
            self.builder.get_object("progressbar1").show()

            count = 0.0
            total = len(data)

            for site in data:                    
                for taxon in data[site]["species_list"]:
                    for code in self.scores:
                        if taxon in self.scores[code]["taxa"]:
                            if code not in data[site]["scoring_species"]:
                                data[site]["scoring_species"].append(code)
                                data[site]["sqs"] = data[site]["sqs"] + self.scores[code]["score"]

                if len(data[site]["scoring_species"]) >= 40:
                    model.append([site, len(data[site]["species_list"]), len(data[site]["scoring_species"]), data[site]["sqs"],round((float(data[site]["sqs"])/float(len(data[site]["scoring_species"])))*100, 1) ])
                else:
                    model.append([site, len(data[site]["species_list"]), len(data[site]["scoring_species"]), data[site]["sqs"],"N/A" ])            

                self.builder.get_object("progressbar1").set_fraction(count/total)
                self.builder.get_object("progressbar1").set_text(''.join(["Processed ", str(int(count)), " of ", str(total), " sites"]))
                count = count + 1.0
                
                while gtk.events_pending():
                    gtk.main_iteration()

        self.builder.get_object("progressbar1").hide()
        self.builder.get_object("window1").window.set_cursor(None)
        self.builder.get_object("vbox1").set_sensitive(True)
        
        while gtk.events_pending():
            gtk.main_iteration()
                
    def main_quit(self, widget, var=None):
        gtk.main_quit()

    def show_about_dialog(self, widget):
       about=gtk.AboutDialog()
       about.set_name("xylic")
       about.set_copyright("2010 Charlie Barnes")
       about.set_authors(["Charlie Barnes <charlie@cucaera.co.uk>"])
       about.set_license("xylic is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the Licence, or (at your option) any later version.\n\nxylic is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.\n\nYou should have received a copy of the GNU General Public License along with xylic; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA")
       about.set_wrap_license(True)
       about.set_website("http://cucaera.co.uk/software/xylic/")
       about.set_transient_for(self.builder.get_object("window1"))
       result=about.run()
       about.destroy()

if __name__ == '__main__':
    xylicActions()
    gtk.main()
    
