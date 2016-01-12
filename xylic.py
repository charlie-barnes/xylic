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

        self.scores = { 1 : { "score" : 2, "taxa" : ["Anaspis (Anaspis) humeralis","Anaspis humeralis"] },
2 : { "score" : 16, "taxa" : ["Anaspis melanostoma","Anaspis melanostoma"] },
3 : { "score" : 16, "taxa" : ["Atomaria (Atomaria) procerula","Atomaria procerula"] },
4 : { "score" : 24, "taxa" : ["Bolitochara reyi","Bolitochara reyi"] },
5 : { "score" : 8, "taxa" : ["Cryptophagus (Cryptophagus) acuminatus","Cryptophagus acuminatus"] },
6 : { "score" : 8, "taxa" : ["Cryptophagus (Cryptophagus) angustus","Cryptophagus angustus"] },
7 : { "score" : 2, "taxa" : ["Phloeopora bernhaueri","Phloeopora bernhaueri (= teres)"] },
8 : { "score" : 24, "taxa" : ["Rhopalodontus perforatus","Rhopalodontus perforatus"] },
9 : { "score" : 32, "taxa" : ["Abdera affinis","Abdera affinis","Carida affinis"] },
10 : { "score" : 8, "taxa" : ["Abdera biflexuosa","Abdera biflexuosa"] },
11 : { "score" : 8, "taxa" : ["Abdera flexuosa","Abdera flexuosa"] },
12 : { "score" : 16, "taxa" : ["Abdera quadrifasciata","Abdera quadrifasciata"] },
13 : { "score" : 16, "taxa" : ["Abdera triguttata","Abdera triguttata"] },
14 : { "score" : 8, "taxa" : ["Abraeus granulum","Abraeus granulum"] },
15 : { "score" : 4, "taxa" : ["Abraeus perpusillus","Abraeus globosus","Abraeus globosus"] },
16 : { "score" : 2, "taxa" : ["Acalles misellus","Acalles misellus (= turbatus)","Acalles turbatus"] },
17 : { "score" : 8, "taxa" : ["Acanthocinus aedilis","Acanthocinus aedilis"] },
18 : { "score" : 24, "taxa" : ["Acritus homoeopathicus","Acritus homoeopathicus"] },
19 : { "score" : 2, "taxa" : ["Acrulia inflata","Acrulia inflata"] },
20 : { "score" : 8, "taxa" : ["Aderus populneus","Aderus populneus"] },
21 : { "score" : 16, "taxa" : ["Aeletes atomarius","Aeletes atomarius"] },
22 : { "score" : 2, "taxa" : ["Agaricochara latissima","Gyrophaena latissima"] },
23 : { "score" : 16, "taxa" : ["Agathidium (Agathidium) pisanum","Agathidium pisanum (=badium)"] },
24 : { "score" : 2, "taxa" : ["Agathidium (Agathidium) seminulum","Agathidium seminulum"] },
25 : { "score" : 16, "taxa" : ["Agathidium (Cyphoceble) arcticum","Agathidium arcticum"] },
26 : { "score" : 2, "taxa" : ["Agathidium (Cyphoceble) nigrinum","Agathidium nigrinum"] },
27 : { "score" : 24, "taxa" : ["Agathidium (Neoceble) confusum","Agathidium confusum"] },
28 : { "score" : 2, "taxa" : ["Agathidium (Neoceble) nigripenne","Agathidium nigripenne"] },
29 : { "score" : 2, "taxa" : ["Agathidium (Neoceble) rotundatum","Agathidium rotundatum"] },
30 : { "score" : 2, "taxa" : ["Agathidium (Neoceble) varians","Agathidium varians"] },
31 : { "score" : 24, "taxa" : ["Agrilus (Agrilus) viridis","Agrilus viridis"] },
32 : { "score" : 8, "taxa" : ["Agrilus (Anambus) angustulus","Agrilus angustulus"] },
33 : { "score" : 8, "taxa" : ["Agrilus (Anambus) biguttatus","Agrilus pannonicus"] },
34 : { "score" : 8, "taxa" : ["Agrilus (Anambus) laticornis","Agrilus laticornis"] },
35 : { "score" : 4, "taxa" : ["Agrilus (Anambus) sinuatus","Agrilus sinuatus"] },
36 : { "score" : 2, "taxa" : ["Alaobia subglabra","Atheta subglabra"] },
37 : { "score" : 2, "taxa" : ["Alosterna tabacicolor","Alosterna tabacicolor"] },
38 : { "score" : 24, "taxa" : ["Amarochara bonnairei","Amarochara bonnairei"] },
39 : { "score" : 2, "taxa" : ["Ampedus balteatus","Ampedus balteatus"] },
40 : { "score" : 32, "taxa" : ["Ampedus cardinalis","Ampedus cardinalis"] },
41 : { "score" : 16, "taxa" : ["Ampedus cinnabarinus","Ampedus cinnabarinus"] },
42 : { "score" : 8, "taxa" : ["Ampedus elongantulus","Ampedus elongatulus"] },
43 : { "score" : 32, "taxa" : ["Ampedus nigerrimus","Ampedus nigerrimus"] },
44 : { "score" : 8, "taxa" : ["Ampedus nigrinus","Ampedus nigrinus"] },
45 : { "score" : 8, "taxa" : ["Ampedus pomorum","Ampedus pomorum"] },
46 : { "score" : 8, "taxa" : ["Ampedus quercicola","Ampedus quercicola (= pomonae)"] },
47 : { "score" : 24, "taxa" : ["Ampedus rufipennis","Ampedus rufipennis"] },
48 : { "score" : 32, "taxa" : ["Ampedus sanguineus","Ampedus sanguineus"] },
49 : { "score" : 16, "taxa" : ["Ampedus sanguinolentus","Ampedus sanguinolentus"] },
50 : { "score" : 32, "taxa" : ["Ampedus tristis","Ampedus tristis"] },
51 : { "score" : 4, "taxa" : ["Anaglyptus mysticus","Anaglyptus mysticus"] },
52 : { "score" : 16, "taxa" : ["Anaspis (Anaspis) bohemica","Anaspis bohemica"] },
53 : { "score" : 1, "taxa" : ["Anaspis (Anaspis) frontalis","Anaspis frontalis"] },
54 : { "score" : 2, "taxa" : ["Anaspis (Anaspis) lurida","Anaspis lurida"] },
55 : { "score" : 1, "taxa" : ["Anaspis (Anaspis) pulicaria","Anaspis pulicaria"] },
56 : { "score" : 24, "taxa" : ["Anaspis (Anaspis) thoracica","Anaspis septentrionalis (= schilskyana)"] },
57 : { "score" : 8, "taxa" : ["Anaspis (Anaspis) thoracica","Anaspis thoracica"] },
58 : { "score" : 2, "taxa" : ["Anaspis (Nassipa) costai","Anaspis costai"] },
59 : { "score" : 1, "taxa" : ["Anaspis (Nassipa) rufilabris","Anaspis rufilabris"] },
60 : { "score" : 24, "taxa" : ["Anastrangalia sanguinolenta","Anastrangalia (= Leptura) sanguinolenta"] },
61 : { "score" : 2, "taxa" : ["Anisotoma castanea","Anisotoma castanea"] },
62 : { "score" : 2, "taxa" : ["Anisotoma glabra","Anisotoma glabra"] },
63 : { "score" : 2, "taxa" : ["Anisotoma humeralis","Anisotoma humeralis"] },
64 : { "score" : 2, "taxa" : ["Anisotoma orbicularis","Anisotoma orbicularis"] },
65 : { "score" : 16, "taxa" : ["Anisoxya fuscula","Anisoxya fuscula"] },
66 : { "score" : 8, "taxa" : ["Anitys rubens","Anitys rubens"] },
67 : { "score" : 1, "taxa" : ["Anobium fulvicorne","Hemicoelus fulvicornis"] },
68 : { "score" : 8, "taxa" : ["Anobium inexspectatum","Anobium inexspectatum"] },
69 : { "score" : 24, "taxa" : ["Anobium nitidum","Hemicoelus nitidus"] },
70 : { "score" : 1, "taxa" : ["Anobium punctatum","Anobium punctatum"] },
71 : { "score" : 2, "taxa" : ["Anomognathus cuspidatus","Anomognathus cuspidatus"] },
72 : { "score" : 24, "taxa" : ["Anoplodera sexguttata","Anoplodera (= Leptura) sexguttata"] },
73 : { "score" : 32, "taxa" : ["Anthaxia (Anthaxia) nitidula","Anthaxia nitidula"] },
74 : { "score" : 4, "taxa" : ["Anthocomus fasciatus","Anthocomus fasciatus"] },
75 : { "score" : 8, "taxa" : ["Aplocnemus impressus","Aplocnemus impressus (=pini)"] },
76 : { "score" : 16, "taxa" : ["Aplocnemus nigricornis","Aplocnemus nigricornis"] },
77 : { "score" : 2, "taxa" : ["Arhopalus rusticus","Arhopalus rusticus"] },
78 : { "score" : 8, "taxa" : ["Aromia moschata","Aromia moschata"] },
79 : { "score" : 2, "taxa" : ["Asemum striatum","Asemum striatum"] },
80 : { "score" : 2, "taxa" : ["Aspidiphorus orbiculatus","Aspidiphorus orbiculatus"] },
81 : { "score" : 16, "taxa" : ["Atheta autumnalis","Atheta autumnalis"] },
82 : { "score" : 16, "taxa" : ["Atheta boletophila","Atheta boletophila"] },
83 : { "score" : 2, "taxa" : ["Atheta liturata","Atheta liturata"] },
84 : { "score" : 16, "taxa" : ["Atomaria (Anchicera) morio","Atomaria morio"] },
85 : { "score" : 24, "taxa" : ["Atomaria (Atomaria) badia","Atomaria badia"] },
86 : { "score" : 16, "taxa" : ["Atomaria (Atomaria) lohsei","Atomaria lohsei"] },
87 : { "score" : 2, "taxa" : ["Atomaria (Atomaria) pulchra","Atomaria pulchra"] },
88 : { "score" : 16, "taxa" : ["Atomaria (Atomaria) puncticollis","Atomaria puncticollis"] },
89 : { "score" : 1, "taxa" : ["Atrecus affinis","Atrecus affinis"] },
90 : { "score" : 16, "taxa" : ["Aulonium trisulcus","Aulonium trisulcum"] },
91 : { "score" : 24, "taxa" : ["Aulonothroscus brevicollis","Aulonothroscus brevicollis"] },
92 : { "score" : 4, "taxa" : ["Axinotarsus ruficollis","Axinotarsus ruficollis"] },
93 : { "score" : 32, "taxa" : ["Batrisodes adnexus","Batrisodes adnexus (=buqueti)"] },
94 : { "score" : 32, "taxa" : ["Batrisodes delaporti","Batrisodes delaporti"] },
95 : { "score" : 8, "taxa" : ["Batrisodes venustus","Batrisodes venustus"] },
96 : { "score" : 2, "taxa" : ["Bibloporus bicolor","Bibloporus bicolor"] },
97 : { "score" : 8, "taxa" : ["Bibloporus minutus","Bibloporus minutus"] },
98 : { "score" : 4, "taxa" : ["Biphyllus lunatus","Biphyllus lunatus"] },
99 : { "score" : 2, "taxa" : ["Bisnius subuliformis","Philonthus subuliformis"] },
100 : { "score" : 4, "taxa" : ["Bitoma crenata","Bitoma crenata"] },
101 : { "score" : 2, "taxa" : ["Bolitochara lucida","Bolitochara lucida"] },
102 : { "score" : 8, "taxa" : ["Bolitochara mulsanti","Bolitochara mulsanti"] },
103 : { "score" : 8, "taxa" : ["Bolitochara pulchra","Bolitochara pulchra"] },
104 : { "score" : 16, "taxa" : ["Bolitophagus reticulatus","Bolitophagus reticulatus"] },
105 : { "score" : 32, "taxa" : ["Bostrichus capucinus","Bostrichus capucinus"] },
106 : { "score" : 32, "taxa" : ["Brachygonus ruficeps","Brachygonus (= Ampedus) ruficeps"] },
107 : { "score" : 16, "taxa" : ["Cadaverota hansseni","Atheta hansseni"] },
108 : { "score" : 4, "taxa" : ["Caenoscelis sibirica","Caenoscelis sibirica"] },
109 : { "score" : 8, "taxa" : ["Calambus bipustulatus","Calambus (= Selatosomus) bipustulatus"] },
110 : { "score" : 32, "taxa" : ["Carabus (Chaetocarabus) intricatus","Carabus intricatus"] },
111 : { "score" : 32, "taxa" : ["Cardiophorus gramineus","Cardiophorus gramineus"] },
112 : { "score" : 32, "taxa" : ["Cardiophorus ruficollis","Cardiophorus ruficollis"] },
113 : { "score" : 8, "taxa" : ["Carpophilus sexpustulatus","Carpophilus sexpustulatus"] },
114 : { "score" : 4, "taxa" : ["Cartodere (Cartodere) constricta","Cartodere constricta"] },
115 : { "score" : 8, "taxa" : ["Cerylon fagi","Cerylon fagi"] },
116 : { "score" : 2, "taxa" : ["Cerylon ferrugineum","Cerylon ferrugineum"] },
117 : { "score" : 4, "taxa" : ["Cerylon histeroides","Cerylon histeroides"] },
118 : { "score" : 16, "taxa" : ["Choragus sheppardi","Choragus sheppardi"] },
119 : { "score" : 32, "taxa" : ["Chrysanthia nigricornis","Chrysanthia nigricornis"] },
120 : { "score" : 8, "taxa" : ["Cicones variegatus","Cicones variegata"] },
121 : { "score" : 2, "taxa" : ["Cis bidentatus","Cis bidentatus"] },
122 : { "score" : 1, "taxa" : ["Cis boleti","Cis boleti"] },
123 : { "score" : 24, "taxa" : ["Cis dentatus","Cis dentatus"] },
124 : { "score" : 2, "taxa" : ["Cis fagi","Cis fagi"] },
125 : { "score" : 2, "taxa" : ["Cis festivus","Cis festivus"] },
126 : { "score" : 4, "taxa" : ["Cis hispidus","Cis hispidus"] },
127 : { "score" : 8, "taxa" : ["Cis jacquemartii","Cis jacquemarti"] },
128 : { "score" : 8, "taxa" : ["Cis lineatocribratus","Cis lineatocribratus"] },
129 : { "score" : 4, "taxa" : ["Cis micans","Cis micans"] },
130 : { "score" : 2, "taxa" : ["Cis nitidus","Cis nitidus"] },
131 : { "score" : 4, "taxa" : ["Cis punctulatus","Cis punctulatus"] },
132 : { "score" : 2, "taxa" : ["Cis pygmaeus","Cis pygmaeus"] },
133 : { "score" : 2, "taxa" : ["Cis vestitus","Cis vestitus"] },
134 : { "score" : 2, "taxa" : ["Cis villosulus","Cis setiger"] },
135 : { "score" : 1, "taxa" : ["Clytus arietis","Clytus arietis"] },
136 : { "score" : 16, "taxa" : ["Colydium elongatum","Colydium elongatum"] },
137 : { "score" : 8, "taxa" : ["Conopalpus testaceus","Conopalpus testaceus"] },
138 : { "score" : 8, "taxa" : ["Corticaria alleni","Corticaria alleni"] },
139 : { "score" : 24, "taxa" : ["Corticaria fagi","Corticaria fagi"] },
140 : { "score" : 16, "taxa" : ["Corticaria longicollis","Corticaria longicollis"] },
141 : { "score" : 16, "taxa" : ["Corticaria polypori","Corticaria polypori"] },
142 : { "score" : 8, "taxa" : ["Corticaria rubripes","Corticaria linearis"] },
143 : { "score" : 8, "taxa" : ["Corticeus bicolor","Corticeus bicolor"] },
144 : { "score" : 24, "taxa" : ["Corticeus unicolor","Corticeus unicolor"] },
145 : { "score" : 2, "taxa" : ["Coryphium angusticolle","Coryphium angusticolle"] },
146 : { "score" : 16, "taxa" : ["Cossonus linearis","Cossonus linearis"] },
147 : { "score" : 8, "taxa" : ["Cossonus parallelepipedus","Cossonus parallelepipedus"] },
148 : { "score" : 8, "taxa" : ["Cryptarcha strigata","Cryptarcha strigata"] },
149 : { "score" : 8, "taxa" : ["Cryptarcha undata","Cryptarcha undata"] },
150 : { "score" : 2, "taxa" : ["Cryptolestes duplicatus","Cryptolestes duplicatus"] },
151 : { "score" : 2, "taxa" : ["Cryptolestes ferrugineus","Cryptolestes ferrugineus"] },
152 : { "score" : 16, "taxa" : ["Cryptophagus confusus","Cryptophagus confusus"] },
153 : { "score" : 24, "taxa" : ["Cryptophagus corticinus","Cryptophagus corticinus"] },
154 : { "score" : 1, "taxa" : ["Cryptophagus dentatus","Cryptophagus dentatus"] },
155 : { "score" : 24, "taxa" : ["Cryptophagus falcozi","Cryptophagus falcozi"] },
156 : { "score" : 16, "taxa" : ["Cryptophagus intermedius","Cryptophagus intermedius"] },
157 : { "score" : 8, "taxa" : ["Cryptophagus labilis","Cryptophagus labilis"] },
158 : { "score" : 16, "taxa" : ["Cryptophagus micaceus","Cryptophagus micaceus"] },
159 : { "score" : 8, "taxa" : ["Cryptophagus ruficornis","Cryptophagus ruficornis"] },
160 : { "score" : 4, "taxa" : ["Ctesias serra","Ctesias serra"] },
161 : { "score" : 16, "taxa" : ["Cyanostolus aeneus","Cyanostolus aeneus"] },
162 : { "score" : 4, "taxa" : ["Cyphea curtula","Cyphea curtula"] },
163 : { "score" : 2, "taxa" : ["Dacne bipustulata","Dacne bipustulata"] },
164 : { "score" : 2, "taxa" : ["Dacne rufifrons","Dacne rufifrons"] },
165 : { "score" : 2, "taxa" : ["Dadobia immersa","Dadobia immersa"] },
166 : { "score" : 2, "taxa" : ["Dasytes aeratus","Dasytes aeratus (= aerosus)"] },
167 : { "score" : 16, "taxa" : ["Dasytes niger","Dasytes niger"] },
168 : { "score" : 8, "taxa" : ["Dasytes plumbeus","Dasytes plumbeus"] },
169 : { "score" : 8, "taxa" : ["Dendrophagus crenatus","Dendrophagus crenatus"] },
170 : { "score" : 1, "taxa" : ["Denticollis linearis","Denticollis linearis"] },
171 : { "score" : 8, "taxa" : ["Dexiogyia corticina","Dexiogyia corticina"] },
172 : { "score" : 8, "taxa" : ["Diacanthous undulatus","Harminius undulatus"] },
173 : { "score" : 24, "taxa" : ["Diaperis boleti","Diaperus boleti"] },
174 : { "score" : 16, "taxa" : ["Dictyoptera aurora","Dictyoptera aurora"] },
175 : { "score" : 1, "taxa" : ["Dinaraea aequata","Dinaraea aequata"] },
176 : { "score" : 2, "taxa" : ["Dinaraea linearis","Dinaraea linearis"] },
177 : { "score" : 32, "taxa" : ["Dinoptera collaris","Dinoptera (= Acmaeops) collaris"] },
178 : { "score" : 8, "taxa" : ["Diplocoelus fagi","Diplocoelus fagi"] },
179 : { "score" : 32, "taxa" : ["Dissoleucas niveirostris","Tropideres niveirostris"] },
180 : { "score" : 16, "taxa" : ["Dorcatoma ambjoerni","Dorcatoma ambjourni"] },
181 : { "score" : 4, "taxa" : ["Dorcatoma chrysomelina","Dorcatoma chrysomelina"] },
182 : { "score" : 16, "taxa" : ["Dorcatoma dresdensis","Dorcatoma dresdensis"] },
183 : { "score" : 8, "taxa" : ["Dorcatoma flavicornis","Dorcatoma flavicornis"] },
184 : { "score" : 16, "taxa" : ["Dorcatoma substriata","Dorcatoma serra"] },
185 : { "score" : 2, "taxa" : ["Dorcus parallelipipedus","Dorcus parallelepipedus"] },
186 : { "score" : 2, "taxa" : ["Dropephylla devillei","Dropephylla devillei (= grandiloqua)"] },
187 : { "score" : 8, "taxa" : ["Dropephylla heerii","Dropephylla heeri"] },
188 : { "score" : 1, "taxa" : ["Dropephylla ioptera","Dropephylla ioptera"] },
189 : { "score" : 1, "taxa" : ["Dropephylla vilis","Dropephylla vilis"] },
190 : { "score" : 16, "taxa" : ["Dryocoetes alni","Dryocoetinus alni"] },
191 : { "score" : 2, "taxa" : ["Dryocoetes autographus","Dryocoetes autographus"] },
192 : { "score" : 2, "taxa" : ["Dryocoetes villosus","Dryocoetinus villosus"] },
193 : { "score" : 2, "taxa" : ["Dryophilus pusillus","Dryophilus pusillus"] },
194 : { "score" : 32, "taxa" : ["Dryophthorus corticalis","Dryophthorus corticalis"] },
195 : { "score" : 32, "taxa" : ["Elater ferrugineus","Elater ferrugineus"] },
196 : { "score" : 4, "taxa" : ["Eledona agricola","Eledona agricola"] },
197 : { "score" : 2, "taxa" : ["Endomychus coccineus","Endomychus coccineus"] },
198 : { "score" : 32, "taxa" : ["Endophloeus markovichianus","Endophloeus markovichianus"] },
199 : { "score" : 32, "taxa" : ["Enedreytes sepicola","Tropideres sepicola"] },
200 : { "score" : 8, "taxa" : ["Enicmus brevicornis","Enicmus brevicornis"] },
201 : { "score" : 8, "taxa" : ["Enicmus fungicola","Enicmus fungicola"] },
202 : { "score" : 8, "taxa" : ["Enicmus rugosus","Enicmus rugosus"] },
203 : { "score" : 2, "taxa" : ["Enicmus testaceus","Emicmus testaceus"] },
204 : { "score" : 2, "taxa" : ["Ennearthron cornutum","Ennearthron cornutum"] },
205 : { "score" : 16, "taxa" : ["Epierus comptus","Epierus comptus"] },
206 : { "score" : 8, "taxa" : ["Epiphanis cornutus","Epiphanis cornutus"] },
207 : { "score" : 8, "taxa" : ["Epuraea (Epuraea) angustula","Epuraea angustula"] },
208 : { "score" : 2, "taxa" : ["Epuraea (Epuraea) biguttata","Epuraea biguttata"] },
209 : { "score" : 8, "taxa" : ["Epuraea (Epuraea) distincta","Epuraea distincta"] },
210 : { "score" : 8, "taxa" : ["Epuraea (Epuraea) fuscicollis","Epuraea fuscicollis"] },
211 : { "score" : 8, "taxa" : ["Epuraea (Epuraea) guttata","Epuraea guttata"] },
212 : { "score" : 8, "taxa" : ["Epuraea (Epuraea) longula","Epuraea longula"] },
213 : { "score" : 1, "taxa" : ["Epuraea (Epuraea) marseuli","Epuraea marseuli (= pusilla)"] },
214 : { "score" : 24, "taxa" : ["Epuraea (Epuraea) neglecta","Epuraea neglecta"] },
215 : { "score" : 2, "taxa" : ["Epuraea (Epuraea) pallescens","Epurea pallescens (= florea)"] },
216 : { "score" : 2, "taxa" : ["Epuraea (Epuraea) rufomarginata","Epuraea rufomarginata"] },
217 : { "score" : 1, "taxa" : ["Epuraea (Epuraea) silacea","Epuraea silacea (= deleta)"] },
218 : { "score" : 8, "taxa" : ["Epuraea (Epuraea) terminalis","Epuraea terminalis (= adumbrata)"] },
219 : { "score" : 8, "taxa" : ["Epuraea (Epuraea) thoracica","Epuraea thoracica"] },
220 : { "score" : 16, "taxa" : ["Epuraea (Epuraea) variegata","Epuraea variegata"] },
221 : { "score" : 2, "taxa" : ["Epuraea (Epuraeanella) limbata","Epuraea limbata"] },
222 : { "score" : 2, "taxa" : ["Ernobius mollis","Ernobius mollis"] },
223 : { "score" : 2, "taxa" : ["Ernobius nigrinus","Ernobius nigrinus"] },
224 : { "score" : 16, "taxa" : ["Ernoporicus caucasicus","Ernoporus caucasicus"] },
225 : { "score" : 8, "taxa" : ["Ernoporicus fagi","Ernoporus fagi"] },
226 : { "score" : 32, "taxa" : ["Ernoporus tiliae","Ernoporus tiliae"] },
227 : { "score" : 32, "taxa" : ["Eucnemis capucina","Eucnemis capucina"] },
228 : { "score" : 32, "taxa" : ["Euconnus (Napochus) pragensis","Euconnus pragensis"] },
229 : { "score" : 8, "taxa" : ["Euglenes oculatus","Aderus oculatus"] },
230 : { "score" : 16, "taxa" : ["Euplectus bescidicus","Euplectus bescidicus"] },
231 : { "score" : 8, "taxa" : ["Euplectus bonvouloiri","Euplectus bonvouloiri"] },
232 : { "score" : 32, "taxa" : ["Euplectus brunneus","Euplectus brunneus"] },
233 : { "score" : 2, "taxa" : ["Euplectus infirmus","Euplectus infirmus"] },
234 : { "score" : 2, "taxa" : ["Euplectus karstenii","Euplectus karsteni"] },
235 : { "score" : 8, "taxa" : ["Euplectus kirbii","Euplectus kirbyi"] },
236 : { "score" : 8, "taxa" : ["Euplectus mutator","Euplectus fauveli"] },
237 : { "score" : 24, "taxa" : ["Euplectus nanus","Euplectus nanus"] },
238 : { "score" : 2, "taxa" : ["Euplectus piceus","Euplectus piceus"] },
239 : { "score" : 24, "taxa" : ["Euplectus punctatus","Euplectus punctatus"] },
240 : { "score" : 24, "taxa" : ["Euryusa optabilis","Euryusa optabilis"] },
241 : { "score" : 24, "taxa" : ["Euryusa sinuata","Euryusa sinuata"] },
242 : { "score" : 32, "taxa" : ["Eutheia formicetorum","Eutheia formicetorum"] },
243 : { "score" : 32, "taxa" : ["Eutheia linearis","Eutheia linearis"] },
244 : { "score" : 1, "taxa" : ["Gabrius splendidulus","Gabrius splendidulus"] },
245 : { "score" : 32, "taxa" : ["Gastrallus immarginatus","Gastrallus immarginatus"] },
246 : { "score" : 16, "taxa" : ["Glaphyra umbellatarum","Molorchus umbellatarum"] },
247 : { "score" : 2, "taxa" : ["Glischrochilus (Glischrochilus) quadripunctatus","Glischrochilus quadripunctatus"] },
248 : { "score" : 2, "taxa" : ["Glischrochilus (Librodor) quadriguttatus","Glischrochilus quadriguttatus"] },
249 : { "score" : 32, "taxa" : ["Globicornis nigripes","Globicornis rufitarsis (=nigripes)"] },
250 : { "score" : 32, "taxa" : ["Gnorimus nobilis","Gnorimus nobilis"] },
251 : { "score" : 32, "taxa" : ["Gnorimus variabilis","Gnorimus variabilis"] },
252 : { "score" : 2, "taxa" : ["Gonodera luperus","Gonodera luperus"] },
253 : { "score" : 24, "taxa" : ["Grammoptera abdominalis","Grammoptera variegata"] },
254 : { "score" : 1, "taxa" : ["Grammoptera ruficornis","Grammoptera ruficornis"] },
255 : { "score" : 24, "taxa" : ["Grammoptera ustulata","Grammoptera ustulata"] },
256 : { "score" : 2, "taxa" : ["Grynobius planus","Grynobius planus"] },
257 : { "score" : 2, "taxa" : ["Gyrophaena bihamata","Gyrophaena bihamata"] },
258 : { "score" : 8, "taxa" : ["Gyrophaena congrua","Gyrophaena congrua"] },
259 : { "score" : 8, "taxa" : ["Gyrophaena joyi","Gyrophaena joyi"] },
260 : { "score" : 8, "taxa" : ["Gyrophaena lucidula","Gyrophaena lucidula"] },
261 : { "score" : 8, "taxa" : ["Gyrophaena manca","Gyrophaena angustata"] },
262 : { "score" : 2, "taxa" : ["Gyrophaena minima","Gyrophaena minima"] },
263 : { "score" : 16, "taxa" : ["Gyrophaena munsteri","Gyrophaena munsteri"] },
264 : { "score" : 16, "taxa" : ["Gyrophaena poweri","Gyrophaena poweri"] },
265 : { "score" : 24, "taxa" : ["Gyrophaena pseudonana","Gyrophaena pseudonana"] },
266 : { "score" : 16, "taxa" : ["Gyrophaena pulchella","Gyrophaena pulchella"] },
267 : { "score" : 8, "taxa" : ["Gyrophaena strictula","Gyrophaena strictula"] },
268 : { "score" : 8, "taxa" : ["Hadrobregmus denticollis","Hadrobregmus denticollis"] },
269 : { "score" : 8, "taxa" : ["Hallomenus binotatus","Hallomenus binotatus"] },
270 : { "score" : 2, "taxa" : ["Hapalaraea pygmaea","Hapalaraea pygmaea"] },
271 : { "score" : 2, "taxa" : ["Haploglossa gentilis","Haploglossa gentilis"] },
272 : { "score" : 8, "taxa" : ["Haploglossa marginalis","Haploglossa marginalis"] },
273 : { "score" : 8, "taxa" : ["Hedobia (Ptinomorphus) imperialis","Ptinomorphus (= Hedobia) imperialis"] },
274 : { "score" : 8, "taxa" : ["Helops caeruleus","Helops caeruleus"] },
275 : { "score" : 2, "taxa" : ["Henoticus serratus","Henoticus serratus"] },
276 : { "score" : 2, "taxa" : ["Homalota plana","Homalota plana"] },
277 : { "score" : 1, "taxa" : ["Hylastes ater","Hylastes ater"] },
278 : { "score" : 2, "taxa" : ["Hylastes brunneus","Hylastes brunneus"] },
279 : { "score" : 2, "taxa" : ["Hylastes opacus","Hylastes opacus"] },
280 : { "score" : 4, "taxa" : ["Hylecoetus dermestoides","Hylecoetus dermestoides"] },
281 : { "score" : 2, "taxa" : ["Hylesinus crenatus","Hylesinus crenatus"] },
282 : { "score" : 8, "taxa" : ["Hylesinus orni","Hylesinus orni"] },
283 : { "score" : 2, "taxa" : ["Hylesinus toranio","Hylesinus oleiperda"] },
284 : { "score" : 1, "taxa" : ["Hylesinus varius","Hylesinus (= Leperisinus) varius"] },
285 : { "score" : 32, "taxa" : ["Hylis cariniceps","Hylis cariniceps"] },
286 : { "score" : 24, "taxa" : ["Hylis olexai","Hylis olexai"] },
287 : { "score" : 1, "taxa" : ["Hylobius (Callirus) abietis","Hylobius abietis"] },
288 : { "score" : 1, "taxa" : ["Hylurgops palliatus","Hylurgops palliatus"] },
289 : { "score" : 32, "taxa" : ["Hypebaeus flavipes","Hypebaeus flavipes"] },
290 : { "score" : 16, "taxa" : ["Hypnogyra angularis","Xantholinus angularis"] },
291 : { "score" : 16, "taxa" : ["Hypulus quercinus","Hypulus quercinus"] },
292 : { "score" : 2, "taxa" : ["Ips acuminatus","Ips acuminatus"] },
293 : { "score" : 16, "taxa" : ["Ischnodes sanguinicollis","Ischnodes sanguinicollis"] },
294 : { "score" : 16, "taxa" : ["Ischnoglossa obscura","Ischnoglossa obscura"] },
295 : { "score" : 2, "taxa" : ["Ischnoglossa prolixa","Ischnoglossa prolixa"] },
296 : { "score" : 2, "taxa" : ["Ischnoglossa turcica","Ischnoglossa turcica"] },
297 : { "score" : 24, "taxa" : ["Ischnomera caerulea","Ischnomera caerulea"] },
298 : { "score" : 32, "taxa" : ["Ischnomera cinerascens","Ischnomera cinerascens"] },
299 : { "score" : 4, "taxa" : ["Ischnomera cyanea","Ischnomera cyanea"] },
300 : { "score" : 8, "taxa" : ["Ischnomera sanguinicollis","Ischnomera sanguinicollis"] },
301 : { "score" : 24, "taxa" : ["Judolia sexmaculata","Judolia sexmaculata"] },
302 : { "score" : 8, "taxa" : ["Kissophagus hederae","Kissophagus hederae"] },
303 : { "score" : 8, "taxa" : ["Korynetes caeruleus","Korynetes caeruleus"] },
304 : { "score" : 8, "taxa" : ["Kyklioacalles roboris","Acalles roboris"] },
305 : { "score" : 32, "taxa" : ["Lacon querceus","Lacon quercus"] },
306 : { "score" : 32, "taxa" : ["Laemophloeus monilis","Laemophloeus monilis"] },
307 : { "score" : 32, "taxa" : ["Lamia textor","Lamia textor"] },
308 : { "score" : 8, "taxa" : ["Latridius consimilis","Lathridius consimilis"] },
309 : { "score" : 2, "taxa" : ["Leiopus nebulosus","Leiopus nebulosus"] },
310 : { "score" : 16, "taxa" : ["Leptura aurulenta","Leptura (= Strangalia) aurulenta"] },
311 : { "score" : 2, "taxa" : ["Leptura quadrifasciata","Leptura (= Strangalia) quadrifasciata"] },
312 : { "score" : 32, "taxa" : ["Lepturobosca virens","Lepturobosca virens"] },
313 : { "score" : 1, "taxa" : ["Leptusa fumida","Leptusa fumida"] },
314 : { "score" : 8, "taxa" : ["Leptusa norvegica","Leptusa norvegica"] },
315 : { "score" : 2, "taxa" : ["Leptusa pulchella","Leptusa pulchella"] },
316 : { "score" : 1, "taxa" : ["Leptusa ruficollis","Leptusa ruficollis"] },
317 : { "score" : 32, "taxa" : ["Limoniscus violaceus","Limoniscus violaceus"] },
318 : { "score" : 16, "taxa" : ["Lissodema cursor","Lissodema cursor"] },
319 : { "score" : 8, "taxa" : ["Lissodema denticolle","Lissodema quadripustulata"] },
320 : { "score" : 2, "taxa" : ["Litargus connexus","Litargus connexus"] },
321 : { "score" : 8, "taxa" : ["Lucanus cervus","Lucanus cervus"] },
322 : { "score" : 4, "taxa" : ["Lyctus brunneus","Lyctus brunneus"] },
323 : { "score" : 8, "taxa" : ["Lyctus linearis","Lyctus linearis"] },
324 : { "score" : 32, "taxa" : ["Lymantor coryli","Lymantor coryli"] },
325 : { "score" : 32, "taxa" : ["Lymexylon navale","Lymexylon navale"] },
326 : { "score" : 2, "taxa" : ["Magdalis (Edo) ruficornis","Magdalis ruficornis"] },
327 : { "score" : 16, "taxa" : ["Magdalis (Magdalis) duplicata","Magdalis duplicata"] },
328 : { "score" : 8, "taxa" : ["Magdalis (Magdalis) phlegmatica","Magdalis phlegmatica"] },
329 : { "score" : 2, "taxa" : ["Magdalis (Odontomagdalis) armigera","Magdalis armigera"] },
330 : { "score" : 4, "taxa" : ["Magdalis (Odontomagdalis) carbonaria","Magdalis carbonaria"] },
331 : { "score" : 8, "taxa" : ["Magdalis (Panus) barbicornis","Magdalis barbicornis"] },
332 : { "score" : 4, "taxa" : ["Magdalis (Porrothus) cerasi","Magdalis cerasi"] },
333 : { "score" : 1, "taxa" : ["Malachius bipustulatus","Malachius bipustulatus"] },
334 : { "score" : 8, "taxa" : ["Malthinus balteatus","Malthinus balteatus"] },
335 : { "score" : 1, "taxa" : ["Malthinus flaveolus","Malthinus flaveolus"] },
336 : { "score" : 8, "taxa" : ["Malthinus frontalis","Malthinus frontalis"] },
337 : { "score" : 2, "taxa" : ["Malthinus seriepunctatus","Malthinus seriepunctatus"] },
338 : { "score" : 24, "taxa" : ["Malthodes crassicornis","Malthodes crassicornis"] },
339 : { "score" : 2, "taxa" : ["Malthodes dispar","Malthodes dispar"] },
340 : { "score" : 8, "taxa" : ["Malthodes fibulatus","Malthodes fibulatus"] },
341 : { "score" : 2, "taxa" : ["Malthodes flavoguttatus","Malthodes flavoguttatus"] },
342 : { "score" : 2, "taxa" : ["Malthodes fuscus","Malthodes fuscus"] },
343 : { "score" : 8, "taxa" : ["Malthodes guttifer","Malthodes guttifer"] },
344 : { "score" : 1, "taxa" : ["Malthodes marginatus","Malthodes marginatus"] },
345 : { "score" : 16, "taxa" : ["Malthodes maurus","Malthodes maurus"] },
346 : { "score" : 1, "taxa" : ["Malthodes minimus","Malthodes minimus"] },
347 : { "score" : 2, "taxa" : ["Malthodes mysticus","Malthodes mysticus"] },
348 : { "score" : 2, "taxa" : ["Malthodes pumilus","Malthodes pumilus"] },
349 : { "score" : 32, "taxa" : ["Megapenthes lugens","Megapenthes lugens"] },
350 : { "score" : 16, "taxa" : ["Megarthrus hemipterus","Megarthrus hemipterus"] },
351 : { "score" : 8, "taxa" : ["Megatoma undata","Megatoma undata"] },
352 : { "score" : 32, "taxa" : ["Melandrya barbata","Melandrya barbata"] },
353 : { "score" : 4, "taxa" : ["Melandrya caraboides","Melandrya caraboides"] },
354 : { "score" : 2, "taxa" : ["Melanophila acuminata","Melanophila acuminata"] },
355 : { "score" : 1, "taxa" : ["Melanotus villosus","Melanotus villosus (= erythropus)"] },
356 : { "score" : 4, "taxa" : ["Melasis buprestoides","Melasis buprestoides"] },
357 : { "score" : 24, "taxa" : ["Mesosa nebulosa","Mesosa nebulosa"] },
358 : { "score" : 16, "taxa" : ["Micrambe (Micrambinus) bimaculata","Micrambe bimaculatus"] },
359 : { "score" : 16, "taxa" : ["Micridium halidaii","Micridium halidaii"] },
360 : { "score" : 8, "taxa" : ["Microrhagus pygmaeus","Microrhagus (= Dirhagus) pygmaeus"] },
361 : { "score" : 24, "taxa" : ["Microscydmus minimus","Microscydmus minimus"] },
362 : { "score" : 16, "taxa" : ["Mordellistena (Mordellistena) neuwaldeggiana","Mordellistena neuwaldeggiana"] },
363 : { "score" : 8, "taxa" : ["Mordellistena (Mordellistena) variegata","Mordellistena variegata"] },
364 : { "score" : 4, "taxa" : ["Mordellochroa abdominalis","Mordellochroa abdominalis"] },
365 : { "score" : 2, "taxa" : ["Mycetaea subterranea","Mycetaea subterranea (= hirta)"] },
366 : { "score" : 16, "taxa" : ["Mycetochara humeralis","Mycetochara humeralis"] },
367 : { "score" : 2, "taxa" : ["Mycetophagus atomarius","Mycetophagus atomarius"] },
368 : { "score" : 32, "taxa" : ["Mycetophagus fulvicollis","Mycetophagus fulvicollis"] },
369 : { "score" : 2, "taxa" : ["Mycetophagus multipunctatus","Mycetophagus multipunctatus"] },
370 : { "score" : 4, "taxa" : ["Mycetophagus piceus","Mycetophagus piceus"] },
371 : { "score" : 16, "taxa" : ["Mycetophagus populi","Mycetophagus populi"] },
372 : { "score" : 16, "taxa" : ["Mycetophagus quadriguttatus","Mycetophagus quadriguttatus"] },
373 : { "score" : 2, "taxa" : ["Mycetophagus quadripustulatus","Mycetophagus quadripustulatus"] },
374 : { "score" : 2, "taxa" : ["Nemadus colonoides","Nemadus colonoides"] },
375 : { "score" : 24, "taxa" : ["Nemozoma elongatum","Nemozoma elongatum"] },
376 : { "score" : 8, "taxa" : ["Neuraphes (Pararaphes) plicicollis","Neuraphes plicicollis"] },
377 : { "score" : 8, "taxa" : ["Nossidium pilosellum","Nossidium pilosellum"] },
378 : { "score" : 16, "taxa" : ["Notolaemus unifasciatus","Notolaemus unifasciatus"] },
379 : { "score" : 2, "taxa" : ["Nudobius lentus","Nudobius lentus"] },
380 : { "score" : 32, "taxa" : ["Obrium cantharinum","Obrium cantharinum"] },
381 : { "score" : 2, "taxa" : ["Ochina ptinoides","Ochina ptinoides"] },
382 : { "score" : 1, "taxa" : ["Octotemnus glabriculus","Octotemnus glabriculus"] },
383 : { "score" : 24, "taxa" : ["Oedemera (Oedemera) virescens","Oedemera virescens"] },
384 : { "score" : 8, "taxa" : ["Oedemera (Oncomera) femoralis","Oncomera femorata"] },
385 : { "score" : 8, "taxa" : ["Opilo mollis","Opilio mollis"] },
386 : { "score" : 4, "taxa" : ["Orchesia micans","Orchesia micans"] },
387 : { "score" : 8, "taxa" : ["Orchesia minor","Orchesia minor"] },
388 : { "score" : 4, "taxa" : ["Orchesia undulata","Orchesia undulata"] },
389 : { "score" : 2, "taxa" : ["Orthocis alni","Cis alni"] },
390 : { "score" : 24, "taxa" : ["Orthocis coluber","Cis coluber"] },
391 : { "score" : 16, "taxa" : ["Orthoperus aequalis","Orthoperus aequalis (= nitidulus)"] },
392 : { "score" : 4, "taxa" : ["Orthoperus nigrescens","Orthoperus mundus"] },
393 : { "score" : 2, "taxa" : ["Orthotomicus suturalis","Orthotomicus suturalis"] },
394 : { "score" : 16, "taxa" : ["Osphya bipunctata","Osphya bipunctata"] },
395 : { "score" : 32, "taxa" : ["Ostoma ferrugineum","Ostoma ferrugineum"] },
396 : { "score" : 32, "taxa" : ["Oxylaemus cylindricus","Oxylaemus cylindricus"] },
397 : { "score" : 24, "taxa" : ["Oxylaemus variolosus","Oxylaemus variolosus"] },
398 : { "score" : 2, "taxa" : ["Pachytodes cerambyciformis","Pachytodes (= Judolia) cerambyciformis"] },
399 : { "score" : 24, "taxa" : ["Paracorymbia fulva","Paracorymbia (= Leptura) fulva"] },
400 : { "score" : 16, "taxa" : ["Paranopleta inhabilis","Paranopleta inhabilis"] },
401 : { "score" : 2, "taxa" : ["Paromalus flavicornis","Paromalus flavicornis"] },
402 : { "score" : 32, "taxa" : ["Paromalus parallelepipedus","Paromalus parallelepipedus"] },
403 : { "score" : 16, "taxa" : ["Pediacus depressus","Pediacus depressus"] },
404 : { "score" : 4, "taxa" : ["Pediacus dermestoides","Pediacus dermestoides"] },
405 : { "score" : 32, "taxa" : ["Pedostrangalia revestita","Pedostrangalia (=Leptura) revestita"] },
406 : { "score" : 4, "taxa" : ["Pentaphyllus testaceus","Pentaphyllus testaceus"] },
407 : { "score" : 2, "taxa" : ["Phloeocharis subtilissima","Phloeocharis subtillissima"] },
408 : { "score" : 2, "taxa" : ["Phloeonomus punctipennis","Phloeonomus punctipennis"] },
409 : { "score" : 2, "taxa" : ["Phloeonomus pusillus","Phloeonomus pusillus"] },
410 : { "score" : 32, "taxa" : ["Phloeophagus gracilis","Phloeophagus (= Rhyncholus) gracilis"] },
411 : { "score" : 2, "taxa" : ["Phloeophagus lignarius","Phloeophagus (= Rhyncholus) lignarius"] },
412 : { "score" : 24, "taxa" : ["Phloeopora concolor","Phloeodroma concolor"] },
413 : { "score" : 8, "taxa" : ["Phloeopora corticalis","Phloeopora corticalis (= angustiformis)"] },
414 : { "score" : 1, "taxa" : ["Phloeopora testacea","Phloeopora testacea"] },
415 : { "score" : 2, "taxa" : ["Phloeostiba lapponica","Phloeostiba lapponica"] },
416 : { "score" : 2, "taxa" : ["Phloeostiba plana","Phloeostiba plana"] },
417 : { "score" : 8, "taxa" : ["Phloiophilus edwardsii","Phloiophilus edwardsi"] },
418 : { "score" : 8, "taxa" : ["Phloiotrya vaudoueri","Phloiotrya vaudoueri"] },
419 : { "score" : 24, "taxa" : ["Phyllodrepa nigra","Phyllodrepa nigra"] },
420 : { "score" : 8, "taxa" : ["Phyllodrepoidea crenata","Phyllodrepoidea crenata"] },
421 : { "score" : 4, "taxa" : ["Phymatodes testaceus","Phymatodes testaceus"] },
422 : { "score" : 2, "taxa" : ["Pissodes castaneus","Pissodes castaneus"] },
423 : { "score" : 2, "taxa" : ["Pissodes pini","Pissodes pini"] },
424 : { "score" : 2, "taxa" : ["Pityogenes bidentatus","Pityogenes bidentatus"] },
425 : { "score" : 16, "taxa" : ["Pityogenes quadridens","Pityogenes quadridens"] },
426 : { "score" : 8, "taxa" : ["Pityogenes trepanatus","Pityogenes trepanatus"] },
427 : { "score" : 2, "taxa" : ["Pityophagus ferrugineus","Pityophagus ferrugineus"] },
428 : { "score" : 24, "taxa" : ["Pityophthorus lichtensteinii","Pityophthorus lichtensteini"] },
429 : { "score" : 2, "taxa" : ["Pityophthorus pubescens","Pityophthorus pubescens"] },
430 : { "score" : 8, "taxa" : ["Placusa depressa","Placusa depressa"] },
431 : { "score" : 2, "taxa" : ["Placusa pumilio","Placusa pumilio"] },
432 : { "score" : 8, "taxa" : ["Placusa tachyporoides","Placusa tachyporoides"] },
433 : { "score" : 32, "taxa" : ["Plagionotus arcuatus","Plagionatus arcuatus"] },
434 : { "score" : 24, "taxa" : ["Platycis cosnardi","Platycis cosnardi"] },
435 : { "score" : 8, "taxa" : ["Platycis minutus","Platycis minutus"] },
436 : { "score" : 32, "taxa" : ["Platydema violaceum","Platydema violaceum"] },
437 : { "score" : 8, "taxa" : ["Platypus cylindrus","Platypus cylindrus"] },
438 : { "score" : 4, "taxa" : ["Platyrhinus resinosus","Platyrhinus resinosus"] },
439 : { "score" : 8, "taxa" : ["Platystomos albinus","Platystomos albinus"] },
440 : { "score" : 32, "taxa" : ["Plectophloeus nitidus","Plectophloeus nitidus"] },
441 : { "score" : 8, "taxa" : ["Plegaderus dissectus","Plegaderus dissectus"] },
442 : { "score" : 16, "taxa" : ["Poecilium alni","Phymatodes alni"] },
443 : { "score" : 16, "taxa" : ["Pogonocherus fasciculatus","Pogonocherus fasciculatus"] },
444 : { "score" : 2, "taxa" : ["Pogonocherus hispidulus","Pogonocherus hispidulus"] },
445 : { "score" : 2, "taxa" : ["Pogonocherus hispidus","Pogonocherus hispidus"] },
446 : { "score" : 8, "taxa" : ["Prionocyphon serricornis","Prionocyphon serricornis"] },
447 : { "score" : 16, "taxa" : ["Prionus coriarius","Prionus coriarius"] },
448 : { "score" : 8, "taxa" : ["Prionychus ater","Prionychus ater"] },
449 : { "score" : 32, "taxa" : ["Prionychus melanarius","Prionychus melanarius"] },
450 : { "score" : 16, "taxa" : ["Procraerus tibialis","Procraerus tibialis"] },
451 : { "score" : 8, "taxa" : ["Pseudocistela ceramboides","Pseudocistela ceramboides"] },
452 : { "score" : 2, "taxa" : ["Pseudophloeophagus aeneopiceus","Caulotrupodes aeneopiceus"] },
453 : { "score" : 4, "taxa" : ["Pseudotriphyllus suturalis","Pseudotriphyllus suturalis"] },
454 : { "score" : 2, "taxa" : ["Pteleobius vittatus","Pteleobius vittatus"] },
455 : { "score" : 8, "taxa" : ["Ptenidium (Gressnerium) gressneri","Ptenidium gressneri"] },
456 : { "score" : 16, "taxa" : ["Ptenidium (Matthewsium) turgidum","Ptenidium turgidum"] },
457 : { "score" : 2, "taxa" : ["Pteryx suturalis","Pteryx suturalis"] },
458 : { "score" : 1, "taxa" : ["Ptilinus pectinicornis","Ptilinus pectinicornis"] },
459 : { "score" : 16, "taxa" : ["Ptiliolum (Euptilium) caledonicum","Ptiliolum caledonicum"] },
460 : { "score" : 2, "taxa" : ["Ptinella aptera","Ptinella aptera"] },
461 : { "score" : 8, "taxa" : ["Ptinella denticollis","Ptinella denticollis"] },
462 : { "score" : 16, "taxa" : ["Ptinella limbata","Ptinella limbata"] },
463 : { "score" : 24, "taxa" : ["Ptinus lichenum","Ptinus lichenum"] },
464 : { "score" : 16, "taxa" : ["Ptinus palliatus","Ptinus palliatus"] },
465 : { "score" : 8, "taxa" : ["Ptinus subpilosus","Ptinus subpilosus"] },
466 : { "score" : 4, "taxa" : ["Pyrochroa coccinea","Pyrochroa coccinea"] },
467 : { "score" : 1, "taxa" : ["Pyrochroa serraticornis","Pyrochroa serraticornis"] },
468 : { "score" : 16, "taxa" : ["Pyropterus nigroruber","Pyropterus nigroruber"] },
469 : { "score" : 24, "taxa" : ["Pyrrhidium sanguineum","Pyrrhidium sanguineum"] },
470 : { "score" : 16, "taxa" : ["Pytho depressus","Pytho depressus"] },
471 : { "score" : 16, "taxa" : ["Quedius (Microsaurus) aetolicus","Quedius aetolicus"] },
472 : { "score" : 8, "taxa" : ["Quedius (Microsaurus) brevicornis","Quedius brevicornis"] },
473 : { "score" : 4, "taxa" : ["Quedius (Microsaurus) maurus","Quedius maurus"] },
474 : { "score" : 8, "taxa" : ["Quedius (Microsaurus) microps","Quedius microps"] },
475 : { "score" : 8, "taxa" : ["Quedius (Microsaurus) scitus","Quedius scitus"] },
476 : { "score" : 8, "taxa" : ["Quedius (Microsaurus) truncicola","Quedius truncicola (=ventralis)"] },
477 : { "score" : 4, "taxa" : ["Quedius (Microsaurus) xanthopus","Quedius xanthopus"] },
478 : { "score" : 2, "taxa" : ["Quedius (Quedionuchus) plagiatus","Quedius plagiatus"] },
479 : { "score" : 16, "taxa" : ["Rabocerus foveolatus","Rabocerus foveolatus"] },
480 : { "score" : 8, "taxa" : ["Rabocerus gabrieli","Rabocerus gabrieli"] },
481 : { "score" : 1, "taxa" : ["Rhagium (Hagrium) bifasciatum","Rhagium bifasciatum"] },
482 : { "score" : 1, "taxa" : ["Rhagium (Megarhagium) mordax","Rhagium mordax"] },
483 : { "score" : 8, "taxa" : ["Rhagium (Rhagium) inquisitor","Rhagium inquisitor"] },
484 : { "score" : 2, "taxa" : ["Rhizophagus (Anomophagus) cribratus","Rhizophagus cribratus"] },
485 : { "score" : 2, "taxa" : ["Rhizophagus (Eurhizophagus) depressus","Rhizophagus depressus"] },
486 : { "score" : 1, "taxa" : ["Rhizophagus (Rhizophagus) bipustulatus","Rhizophagus bipustulatus"] },
487 : { "score" : 1, "taxa" : ["Rhizophagus (Rhizophagus) dispar","Rhizophagus dispar"] },
488 : { "score" : 2, "taxa" : ["Rhizophagus (Rhizophagus) ferrugineus","Rhizophagus ferrugineus"] },
489 : { "score" : 4, "taxa" : ["Rhizophagus (Rhizophagus) nitidulus","Rhizophagus nitidulus"] },
490 : { "score" : 24, "taxa" : ["Rhizophagus (Rhizophagus) oblongicollis","Rhizophagus oblongicollis"] },
491 : { "score" : 2, "taxa" : ["Rhizophagus (Rhizophagus) parallelocollis","Rhizophagus parallelocollis"] },
492 : { "score" : 24, "taxa" : ["Rhizophagus (Rhizophagus) parvulus","Rhizophagus parvulus"] },
493 : { "score" : 2, "taxa" : ["Rhizophagus (Rhizophagus) perforatus","Rhizophagus perforatus"] },
494 : { "score" : 16, "taxa" : ["Rhizophagus (Rhizophagus) picipes","Rhizophagus picipes"] },
495 : { "score" : 8, "taxa" : ["Rhopalomesites tardyi","Mesites tardii"] },
496 : { "score" : 8, "taxa" : ["Rhyncolus ater","Rhyncolus chloropus (=Eremotes ater)"] },
497 : { "score" : 1, "taxa" : ["Rutpela maculata","Rutpela (= Strangalia) maculata"] },
498 : { "score" : 1, "taxa" : ["Salpingus planirostris","Rhinosimus planirostris"] },
499 : { "score" : 1, "taxa" : ["Salpingus ruficollis","Rhinosimus ruficollis"] },
500 : { "score" : 16, "taxa" : ["Saperda carcharias","Saperda carcharias"] },
501 : { "score" : 8, "taxa" : ["Saperda scalaris","Saperda scalaris"] },
502 : { "score" : 2, "taxa" : ["Scaphidium quadrimaculatum","Scaphidium quadrimaculatum"] },
503 : { "score" : 2, "taxa" : ["Scaphisoma agaricinum","Scaphisoma agaricinum"] },
504 : { "score" : 24, "taxa" : ["Scaphisoma assimile","Scaphisoma assimile"] },
505 : { "score" : 8, "taxa" : ["Scaphisoma boleti","Scaphisoma boleti"] },
506 : { "score" : 16, "taxa" : ["Schizotus pectinicornis","Schizotus pectinicornis"] },
507 : { "score" : 2, "taxa" : ["Scolytus intricatus","Scolytus intricatus"] },
508 : { "score" : 8, "taxa" : ["Scolytus mali","Scolytus mali"] },
509 : { "score" : 1, "taxa" : ["Scolytus multistriatus","Scolytus multistriatus"] },
510 : { "score" : 8, "taxa" : ["Scolytus ratzeburgi","Scolytus ratzeburgi"] },
511 : { "score" : 2, "taxa" : ["Scolytus rugulosus","Scolytus rugulosus"] },
512 : { "score" : 2, "taxa" : ["Scolytus scolytus","Scolytus scolytus"] },
513 : { "score" : 32, "taxa" : ["Scraptia dubia","Scraptia dubia"] },
514 : { "score" : 32, "taxa" : ["Scraptia fuscula","Scraptia fuscula"] },
515 : { "score" : 16, "taxa" : ["Scraptia testacea","Scraptia testacea"] },
516 : { "score" : 24, "taxa" : ["Scydmaenus (Cholerus) rufus","Scydmaenus rufus"] },
517 : { "score" : 8, "taxa" : ["Sepedophilus bipunctatus","Sepedophilus bipunctatus"] },
518 : { "score" : 2, "taxa" : ["Sepedophilus littoreus","Sepedophilus littoreus"] },
519 : { "score" : 2, "taxa" : ["Sepedophilus lusitanicus","Sepedophilus lusitanicus"] },
520 : { "score" : 8, "taxa" : ["Sepedophilus testaceus","Sepedophilus testaceus"] },
521 : { "score" : 2, "taxa" : ["Siagonium quadricorne","Siagonum quadricorne"] },
522 : { "score" : 8, "taxa" : ["Silusa rubiginosa","Silusa rubiginosa"] },
523 : { "score" : 32, "taxa" : ["Silvanoprus fagi","Silvanoprus fagi"] },
524 : { "score" : 8, "taxa" : ["Silvanus bidentatus","Silvanus bidentatus"] },
525 : { "score" : 4, "taxa" : ["Silvanus unidentatus","Silvanus unidentatus"] },
526 : { "score" : 2, "taxa" : ["Sinodendron cylindricum","Sinodendron cylindricum"] },
527 : { "score" : 2, "taxa" : ["Soronia grisea","Soronia grisea"] },
528 : { "score" : 2, "taxa" : ["Soronia punctatissima","Soronia punctatissima"] },
529 : { "score" : 2, "taxa" : ["Sphaeriestes castaneus","Salpingus castaneus"] },
530 : { "score" : 2, "taxa" : ["Sphaeriestes reyi","Salpingus ater"] },
531 : { "score" : 2, "taxa" : ["Sphaeriestes reyi","Salpingus reyi"] },
532 : { "score" : 8, "taxa" : ["Sphindus dubius","Sphindus dubius"] },
533 : { "score" : 16, "taxa" : ["Sphinginus lobatus","Sphinginus lobatus"] },
534 : { "score" : 4, "taxa" : ["Stenagostus rhombeus","Stenagostus rhombeus (= villosus)"] },
535 : { "score" : 4, "taxa" : ["Stenichnus bicolor","Stenichnus bicolor"] },
536 : { "score" : 24, "taxa" : ["Stenichnus godarti","Stenichnus godarti"] },
537 : { "score" : 2, "taxa" : ["Stenocorus meridianus","Stenocorus meridianus"] },
538 : { "score" : 8, "taxa" : ["Stenostola dubia","Stenostola dubia"] },
539 : { "score" : 2, "taxa" : ["Stenurella melanura","Stenurella (= Strangalia) melanura"] },
540 : { "score" : 24, "taxa" : ["Stenurella nigra","Stenurella (= Strangalia) nigra"] },
541 : { "score" : 4, "taxa" : ["Stephostethus alternans","Stephostethus alternans"] },
542 : { "score" : 16, "taxa" : ["Stereocorynes truncorum","Stereocorynes (= Rhyncholus) truncorum"] },
543 : { "score" : 24, "taxa" : ["Stichoglossa semirufa","Stichoglossa semirufa"] },
544 : { "score" : 16, "taxa" : ["Stictoleptura scutellata","Stictoleptura (=Anoplodera) scutellata"] },
545 : { "score" : 8, "taxa" : ["Strigocis bicornis","Sulcacis bicornis"] },
546 : { "score" : 2, "taxa" : ["Sulcacis affinis","Sulcacis affinis"] },
547 : { "score" : 8, "taxa" : ["Symbiotes latus","Symbiotes latus"] },
548 : { "score" : 8, "taxa" : ["Synchita humeralis","Synchita humeralis"] },
549 : { "score" : 24, "taxa" : ["Synchita separanda","Synchita separanda"] },
550 : { "score" : 32, "taxa" : ["Tachinus bipustulatus","Tachinus bipustulatus"] },
551 : { "score" : 32, "taxa" : ["Tachyusida gracilis","Tachyusida gracilis"] },
552 : { "score" : 8, "taxa" : ["Taphrorhychus bicolor","Taphrorhychus bicolor"] },
553 : { "score" : 32, "taxa" : ["Tarsostenus univittatus","Tarsostenus univittatus"] },
554 : { "score" : 32, "taxa" : ["Teredus cylindricus","Teredus cylindricus"] },
555 : { "score" : 32, "taxa" : ["Teretrius fabricii","Teretrius fabricii"] },
556 : { "score" : 8, "taxa" : ["Tetratoma ancora","Tetratoma ancora"] },
557 : { "score" : 16, "taxa" : ["Tetratoma desmarestii","Tetratoma desmaresti"] },
558 : { "score" : 2, "taxa" : ["Tetratoma fungorum","Tetratoma fungorum"] },
559 : { "score" : 2, "taxa" : ["Tetrops praeustus","Tetrops praeusta"] },
560 : { "score" : 16, "taxa" : ["Tetrops starkii","Tetrops starkii"] },
561 : { "score" : 2, "taxa" : ["Thamiaraea cinnamomea","Thamiaraea cinnamomea"] },
562 : { "score" : 8, "taxa" : ["Thamiaraea hospita","Thamiaraea hospita"] },
563 : { "score" : 24, "taxa" : ["Thanasimus femoralis","Thanasimus rufipes"] },
564 : { "score" : 4, "taxa" : ["Thanasimus formicarius","Thanasimus formicarius"] },
565 : { "score" : 8, "taxa" : ["Thymalus limbatus","Thymalus limbatus"] },
566 : { "score" : 32, "taxa" : ["Tilloidea unifasciata","Tilloidea unifasciatus"] },
567 : { "score" : 8, "taxa" : ["Tillus elongatus","Tillus elongatus"] },
568 : { "score" : 24, "taxa" : ["Tomicus minor","Tomicus minor"] },
569 : { "score" : 1, "taxa" : ["Tomicus piniperda","Tomicus piniperda"] },
570 : { "score" : 16, "taxa" : ["Tomoxia bucephala","Tomoxia bucephala (= biguttata)"] },
571 : { "score" : 8, "taxa" : ["Trachodes hispidus","Trachodes hispidus"] },
572 : { "score" : 2, "taxa" : ["Trichius fasciatus","Trichius fasciatus"] },
573 : { "score" : 32, "taxa" : ["Trichonyx sulcicollis","Trichonyx sulcicollis"] },
574 : { "score" : 24, "taxa" : ["Trinodes hirtus","Trinodes hirtus"] },
575 : { "score" : 4, "taxa" : ["Triphyllus bicolor","Triphyllus bicolor"] },
576 : { "score" : 2, "taxa" : ["Triplax aenea","Triplax aenea"] },
577 : { "score" : 24, "taxa" : ["Triplax lacordairii","Triplax lacordairii"] },
578 : { "score" : 4, "taxa" : ["Triplax russica","Triplax russica"] },
579 : { "score" : 32, "taxa" : ["Triplax scutellaris","Triplax scutellaris"] },
580 : { "score" : 16, "taxa" : ["Tritoma bipustulata","Tritoma bipustulata"] },
581 : { "score" : 2, "taxa" : ["Trypodendron domesticum","Trypodendron (= Xyloterus) domesticum"] },
582 : { "score" : 2, "taxa" : ["Trypodendron lineatum","Trypodendron (= Xyloterus) lineatum"] },
583 : { "score" : 8, "taxa" : ["Trypodendron signatum","Trypodendron (= Xyloterus) signatum"] },
584 : { "score" : 16, "taxa" : ["Trypophloeus binodulus","Trypophloeus binodulus (= asperatus)"] },
585 : { "score" : 32, "taxa" : ["Trypophloeus granulatus","Trypophloeus granulatus"] },
586 : { "score" : 16, "taxa" : ["Uleiota planata","Uleiota planata"] },
587 : { "score" : 32, "taxa" : ["Vanonus brevicornis","Aderus brevicornis"] },
588 : { "score" : 32, "taxa" : ["Velleius dilatatus","Velleius dilatatus"] },
589 : { "score" : 2, "taxa" : ["Vincenzellus ruficollis","Vicenzellus ruficollis"] },
590 : { "score" : 4, "taxa" : ["Xestobium rufovillosum","Xestobium rufovillosum"] },
591 : { "score" : 4, "taxa" : ["Xyleborinus saxesenii","Xyleborinus saxeseni"] },
592 : { "score" : 8, "taxa" : ["Xyleborus dispar","Xyleborus dispar"] },
593 : { "score" : 8, "taxa" : ["Xyleborus dryographus","Xyleborus dryographus"] },
594 : { "score" : 32, "taxa" : ["Xyletinus longitarsis","Xyletinus longitarsus"] },
595 : { "score" : 16, "taxa" : ["Xylita laevigata","Xylita laevigata"] },
596 : { "score" : 32, "taxa" : ["Xylodromus testaceus","Xylodromus testaceus"] },
597 : { "score" : 8, "taxa" : ["Xylostiba monilicornis","Xylostiba monilicornis"] },
598 : { "score" : 8, "taxa" : ["Zilora ferruginea","Zilora ferruginea"] },
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
    
