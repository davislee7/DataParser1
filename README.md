# DataParser1
Spreadsheet parser in C#. Designed to be flexible to handle different file types using an interface. Currently checks and validates cell data in .csv and .xlsx file types.

Example Input: davis2.xlsx (or davis2.csv)

statecode	county	line	construction	tiv_2011	tiv_2012	eq_site_deductible	hu_site_deductible	point_latitude	point_longitude
FL	CLAY COUNTY	Residential	Masonry	498960	792148.9	0	9979.2	30.102261	-81.711777
FL	CLAY COUNTY	Residential	Masonry	1322376.3	1438163.57	0	0	30.063936	-81.707664
FL	CLAY COUNTY	Residential	Wood	190724.4	192476.78	0	0	30.089579	-81.700455
FL	CLAY COUNTY	Residential	Wood	79520.76	86854.48	0	0	30.063236	-81.707703
FL	CLAY COUNTY	Residential	Wood	254281.5	246144.49	0	0	30.060614	-81.702675
FL	CLAY COUNTY	Residential	Masonry	515035.62	884419.17	0	0	30.063236	-81.707703
FL	CLAY COUNTY	Commercial	Reinforced Concrete	19260000	20610000	0	0	30.102226	-81.713882
FL	CLAY COUNTY	Residential	Wood	328500	348374.25	0	16425	30.102217	-81.707146
FL	CLAY COUNTY	Residential	Wood	315000	265821.57	0	15750	30.118774	-81.704613
FL	CLAY COUNTY	Residential	Masonry	705600	1010842.56	14112	35280	30.100628	-81.703751
FL	CLAY COUNTY	Residential	Masonry	831498.3	1117791.48	0	0	30.10216	-81.719444
FL	CLAY COUNTY	Residential	Wood	24059.09	33952.19	0	0	30.095957	-81.695099
FL	CLAY COUNTY	Residential	Wood	48115.94	66755.39	0	0	30.100073	-81.739822
FL	CLAY COUNTY	Residential	Wood	28869.12	42826!99	0	0	30.09248	-81.725167
FL	CLAY COUNTY	Residential	Wood	56135.64	50656.8	0	0	30.101356	-81.726248
FL	CLAY COUNTY	Residential	Wood	48115.94	67905.07	0	0	30.113743	-81.727463
FL	CLAY COUNTY	Residential	2947	48115.94	66938.9	0	Etet	30.121655	-81.732391
FL	CLAY COUNTY	Residential	Wood	80192.49	86421.04	0	0	30.109537	-81.741661
FL	CLAY COUNTY	Residential	Wood	48115.94	73798.5	0	{}}	30.11824	-81.745335
FL	CLAY COUNTY	Residential	Wood	60946.79	62467.29	0	0	30.065799	-81.717416
FL	CLAY COUNTY	Residential	Wood	28869.12	42727.74	0	0	30.082993	-81.710581
FL	CLAY COUNTY	Commercial	Reinforced Concrete	13410000	11700000	0	0	30.091921	-81.711929
FL	CLAY COUNTY	Residential	Masonry	1669113.93	2099127.76	0	0	30.117352	-81.711884
FL	CLAY COUNTY	Residential	Wood	179562.23	211372.57	0	0	30.095783	-81.713181
FL	CLAY COUNTY	Residential	Wood	177744.16	157171.16	#NAME?	0	30.110518	-81.727478
FL	CLAY COUNTY	Residential	Wood	17757.58	16948.72	0	0	30.10288	-81.705719
FL	CLAY COUNTY	Residential	Wood	130129.87	101758.43	0	#NAME?	30.068468	-81.71@#=24
FL	CLAY COUNTY	Residential	Wood	42854.77	63592.88	0	0	30.068468	-81.71@#=24
FL	CLAY COUNTY	Residential	Wood	785.58	662.18	0	0	30.068468	-81.71@#=24
FL	CLAY COUNTY	Residential	Wood	170361.91	177176.38	0	0	30.068468	-81.71@#=24
FL	CLAY COUNTY	Residential	Wood	1430.89	1861.41	0	0	30.068468	-81.71@#=24
FL	CLAY COUNTY	Residential	Wood	129913.27	101692.86	0	RFw	30.079785	-81.706865
FL	CLAY COUNTY	Residential	Masonry	366285.62	507164.19	0	0	30.08012	-81.718452
FL	CLAY COUNTY	Residential	Wood	22512.61	28637.17	0	0	30.08012	-81.718452
FL	SUWANNEE COUNTY	Residential	Wood	9246.6	10880.22	0	0	29.959805	-82.926659
FL	SUWANNEE COUNTY	Residential	42r2r4	96164.64	69357.78	0	0	29.959805	-82.926659
FL	SUWANNEE COUNTY	Residential	Wood	11095.92	12737.89	0	0	29.959805	-82.926659
FL	SUWANNEE COUNTY	Residential	Wood	218475	199030.29	0	4369.5	29.9626o1	-82.926155
FL	SUWANNEE COUNTY	Residential	Masonry	1400904	1772984.1	0	0	29.9626o1	-82.926155
FL	SUWANNEE COUNTY	Residential	Wood	4365	4438.05	0	87.3	29.9626o1	-82.926155
FL	SUWANNEE COUNTY	Residential	Wood	4365	6095.72	0	87.3	29.9626o1	-82.926155
FL	SUWANNEE COUNTY	Residential	Wood	39789	58106.58	0	0	29.9626o1	-82.926155
FL	SUWANNEE COUNTY	Residential	Wood	24867	18969.79	0	0	29.9626o1	-82.926155
FL	SUWANNEE COUNTY	Residential	Wood	213876	261435.18	0	0	29.9626o1	-82.926155
FL	SUWANNEE COUNTY	Residential	Wood	69435	93674.34	0	1388.7	29.960735	-82.92542
FL	SUWANNEE COUNTY	Residential	Wood	14922	12333.03	0	0	29.960735	-82.92542
FL	SUWANNEE COUNTY	Residential	Wood	165546	239134.51	0	0	29.963396	-82.916763
FL	SUWANNEE COUNTY	Residential	Wood	72837	86637.86	0	0	29.963396	-82.916763
FL	SUWANNEE COUNTY	Residential	Wood	72837	98147.86	0	0	29.963396	-82.916763
FL	SUWANNEE COUNTY	Residential	Wood	19440	30658.59	0	388.8	29.963396	-82.916763
FL	SUWANNEE COUNTY	Residential	Wood	9945	11551.12	0	198.9	29.963396	-82.916763
FL	SUWANNEE COUNTY	Residential	Wood	255878.61	345894.66	0	0	29.958415	-82.92394
FL	SUWANNEE COUNTY	Residential	Wood	153527.17	138228.8	0	0	29.9696	-82.92767
FL	SUWANNEE COUNTY	Residential	Wood	255878.61	177339.74	0	0	29.959404	-82.927582
FL	SUWANNEE COUNTY	Residential	Wood	102351.45	95132.8	0	0	29.958904	-82.922729
FL	SUWANNEE COUNTY	Residential	Wood	155489.66	145139.65	0	0	29.95822	-82.922424
FL	SUWANNEE COUNTY	Residential	Wood	137233.8	114919.58	0	0	29.965832	-82.933777
FL	SUWANNEE COUNTY	Residential	Wood	123596.05	183015.09	0	0	29.965721	-82.933777
FL	SUWANNEE COUNTY	Residential	Wood	107111.11	88218.85	0	0	29.965717	-82.933777
FL	SUWANNEE COUNTY	Residential	Wood	96309.32	85911	0	0	29.95775	-82.923635
FL	NASSAU COUNTY	Residential	Wood	104031.7	168443.97	0	0	30.39431	-81.93397
FL	NASSAU COUNTY	Residential	Wood	338944.5	485816.61	0	0	30.56267	-81.830429
FL	NASSAU COUNTY	Residential	Wood	272349	414565.56	0	0	30.56267	-81.830429
FL	NASSAU COUNTY	Residential	Wood	129690	129635.53	0	0	30.561651	-81.830193
FL	NASSAU COUNTY	Residential	Wood	123210	120345.61	0	0	30.56267	-81.830429
FL	NASSAU COUNTY	Residential	Wood	3698.64	3939.13	0	0	30.56106	-81.82632
FL	NASSAU COUNTY	Commercial	Reinforced Masonry	2115760.5	3057739.39	0	RTE	30.5579	-81.8249
FL	NASSAU COUNTY	Residential	Wood	93548.58	118401.07	0	0	30.56106	-81.82632
FL	NASSAU COUNTY	Residential	Masonry	1037311.14	1487255.23	0	0	30.56106	-81.82632
FL	NASSAU COUNTY	Residential	Wood	15463.3	11825.56	0	0	30.57185	-81.82383
FL	NASSAU COUNTY	Residential	Wood	13893.35	13988.5	0	0	30.562835	-81.826525
FL	NASSAU COUNTY	Residential	Wood	316618.23	2302!99	0	0	30.567566	-81.826584
FL	NASSAU COUNTY	Residential	Wood	54834.66	63053.72	^%	0	30.60895	-81.83645
FL	NASSAU COUNTY	Residential	Wood	295269.11	488965.65	0	**)	30.59925	-81.81973
FL	NASSAU COUNTY	Residential	Masonry	723734.17	955908.09	0	0	30.53674	-81.77496
FL	NASSAU COUNTY	Residential	Masonry	1741572	1455184.42	0	0	30.555161	-81.825684
FL	COLUMBIA COUNTY	Residential	Wood	135450	164978.1	0	0	30.105968	-82.66227
FL	46575	Residential	Wood	60300	77678.46	0	0	30.1018	-82.7094
FL	COLUMBIA COUNTY	Residential	Wood	22500	19501.76	0	0	30.1018	-82.7094
FL	232223424	Residential	Wood	34212.42	33250.37	0	0	30.1018	-82.7094
FL	COLUMBIA COUNTY	Residential	2r2r2g1 vb 	460128.2	626142.46	0	0	30.1653	-82.73404
FL	COLUMBIA COUNTY	Residential	Wood	32169.69	37291.11	0	0	30.1018	-82.7094
FL	COLUMBIA COUNTY	Residential	Wood	50550.75	49526.38	0	0	30.17399	-82.57017
FL	COLUMBIA COUNTY	Residential	Wood	42300	53762.62	0	0	30.1515	-82.6126
FL	COLUMBIA COUNTY	Residential	Wood	10171.26	16798.29	0	%^%	30.16269	-82.64
FL	COLUMBIA COUNTY	Residential	Wood	155342.88	176189.89	0	0	30.16269	-82.64
FL	COLUMBIA COUNTY	Residential	Wood	3698.64	4424.48	0	0	30.16269	-82.64
FL	COLUMBIA COUNTY	Residential	Wood	4623.3	4525.86	???	0	30.16269	-82.64
FL	COLUMBIA COUNTY	Residential	Wood	33287.76	27375.62	0	0	30.16269	-82.64
FL	COLUMBIA COUNTY	Residential	Wood	21267.18	24840.49	0	0	30.163641	-82.640976
FL	COLUMBIA COUNTY	Residential	Masonry	729556.74	889825.76	0	0	30.163641	-82.640976
FL	112232332	Residential	Wood	350446.14	438049.26	0	0	30.163641	-82.640976
FL	COLUMBIA COUNTY	Commercial	Reinforced Masonry	2135964.6	1728721.59	$#@$	0	30.163641	-82.640976
FL	COLUMBIA COUNTY	Residential	Wood	4623.3	4782.1	0	0	30.163641	-82.640976
FL	COLUMBIA COUNTY	Commercial	Reinforced Masonry	2469639.6	2708463.63	0	0	30.1515	-82.6126
FL	COLUMBIA COUNTY	Residential	Masonry	914073.3	1296439.3	0	!@#$	30.1515	-82.6126
FL	COLUMBIA COUNTY	Commercial	2rt21cgt vy2	3042940.5	3181868!99	)*	0	30.1515	-82.6126
FL	COLUMBIA COUNTY	Residential	Wood	234997.65	255975.89	0	0	30.1515	-82.6126
FL	COLUMBIA COUNTY	Residential	Masonry	912600	1147247.71	0	0	30.1515	-82.6126
FL	COLUMBIA COUNTY	Commercial	Reinforced Masonry	1920744.9	2837762.3	*(	0	30.1515	-82.6126

Example Output: 
