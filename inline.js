
    const LIMIT_QTY = 2000;
    const LIMIT_LOTS = 10;
    const NEAR_QTY = 1800;
    const NEAR_LOTS = 9;
    const PLANILHA_URL = "Painel%20de%20series%20de%20DM.xlsx";
    const EMBEDDED_LOTES = [{"data":"2026-01-05","lote":2344,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 01 - FMT","linha_planilha":2},{"data":"2026-01-05","lote":2345,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 01 - MP","linha_planilha":3},{"data":"2026-01-06","lote":2347,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":318,"serie_planilha":"Série 01 - MP","linha_planilha":4},{"data":"2026-01-06","lote":2348,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 01 - MP","linha_planilha":5},{"data":"2026-01-07","lote":2349,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 01 - FMT","linha_planilha":6},{"data":"2026-01-05","lote":2473,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":100,"serie_planilha":"Série 01 - FN","linha_planilha":7},{"data":"2026-01-06","lote":2474,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":95,"serie_planilha":"Série 01 - FN","linha_planilha":8},{"data":"2026-01-07","lote":2475,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":100,"serie_planilha":"Série 01 - FN","linha_planilha":9},{"data":"2026-01-15","lote":2480,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":100,"serie_planilha":"Série 01 - FN","linha_planilha":10},{"data":"2026-01-16","lote":2481,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":100,"serie_planilha":"Série 03 - FN","linha_planilha":11},{"data":"2026-01-17","lote":2482,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":100,"serie_planilha":"Série 02 - FN","linha_planilha":12},{"data":"2026-01-19","lote":2483,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":100,"serie_planilha":"Série 02 - FN","linha_planilha":13},{"data":"2026-01-20","lote":2485,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":100,"serie_planilha":"Série 02 - FN","linha_planilha":14},{"data":"2026-01-21","lote":2487,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":100,"serie_planilha":"Série 02 - FN","linha_planilha":15},{"data":"2026-01-26","lote":2417,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":165,"serie_planilha":"Série 03 - FN","linha_planilha":16},{"data":"2026-01-27","lote":2418,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":195,"serie_planilha":"Série 03 - FN","linha_planilha":17},{"data":"2026-01-28","lote":2419,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":235,"serie_planilha":"Série 03 - FN","linha_planilha":18},{"data":"2026-01-29","lote":2420,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":255,"serie_planilha":"Série 03 - FN","linha_planilha":19},{"data":"2026-01-07","lote":2501,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 01 - MP","linha_planilha":20},{"data":"2026-01-07","lote":2502,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":318,"serie_planilha":"Série 01 - MP","linha_planilha":21},{"data":"2026-01-08","lote":2503,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 01 - MP","linha_planilha":22},{"data":"2026-01-08","lote":2504,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 01 - FMT","linha_planilha":23},{"data":"2026-01-08","lote":2505,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":318,"serie_planilha":"Série 01 - FMT","linha_planilha":24},{"data":"2026-01-31","lote":2506,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":270,"serie_planilha":"Série 03 - FN","linha_planilha":25},{"data":"2026-02-02","lote":2507,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 03 - FN","linha_planilha":26},{"data":"2026-02-03","lote":2508,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 03 - FN","linha_planilha":27},{"data":"2026-01-09","lote":2509,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":11,"serie_planilha":"Série 02 - MP","linha_planilha":28},{"data":"2026-01-09","lote":2509,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":319,"serie_planilha":"Série 02 - MP","linha_planilha":29},{"data":"2026-01-09","lote":2510,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 01 - FMT","linha_planilha":30},{"data":"2026-01-09","lote":2511,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":318,"serie_planilha":"Série 01 - FMT","linha_planilha":31},{"data":"2026-01-12","lote":2512,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 02 - MP","linha_planilha":32},{"data":"2026-01-12","lote":2513,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 02 - FMT","linha_planilha":33},{"data":"2026-01-12","lote":2514,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":318,"serie_planilha":"Série 02 - FMT","linha_planilha":34},{"data":"2026-01-13","lote":2515,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 02 - MP","linha_planilha":35},{"data":"2026-01-13","lote":2516,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 02 - FMT","linha_planilha":36},{"data":"2026-01-13","lote":2517,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":312,"serie_planilha":"Série 02 - FMT","linha_planilha":37},{"data":"2026-01-14","lote":2518,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 02 - FMT","linha_planilha":38},{"data":"2026-01-14","lote":2519,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 02 - MP","linha_planilha":39},{"data":"2026-01-14","lote":2520,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":312,"serie_planilha":"Série 02 - FMT","linha_planilha":40},{"data":"2026-01-15","lote":2521,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 03 - FMT","linha_planilha":41},{"data":"2026-01-16","lote":2522,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 02 - MP","linha_planilha":42},{"data":"2026-01-15","lote":2523,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":312,"serie_planilha":"Série 03 - FMT","linha_planilha":43},{"data":"2026-01-16","lote":2524,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 03 - FMT","linha_planilha":44},{"data":"2026-01-16","lote":2525,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":312,"serie_planilha":"Série 03 - FMT","linha_planilha":45},{"data":"2026-01-17","lote":2526,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 03 - FMT","linha_planilha":46},{"data":"2026-01-19","lote":2527,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 04 - FMT","linha_planilha":47},{"data":"2026-01-19","lote":2528,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 03 - FMT","linha_planilha":48},{"data":"2026-01-19","lote":2529,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 02 - MP","linha_planilha":49},{"data":"2026-01-20","lote":2530,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 04 - FMT","linha_planilha":50},{"data":"2026-01-20","lote":2531,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 04 - FMT","linha_planilha":51},{"data":"2026-01-20","lote":2532,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 03 - MP","linha_planilha":52},{"data":"2026-01-21","lote":2533,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 04 - FMT","linha_planilha":53},{"data":"2026-01-21","lote":2534,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 04 - FMT","linha_planilha":54},{"data":"2026-01-22","lote":2535,"projeto":"MALHA PAULISTA","tipo":"Bitola Mista MP - USP","quantidade":330,"serie_planilha":"Série 03 - MP","linha_planilha":55},{"data":"2026-01-22","lote":2536,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 05 - FMT","linha_planilha":56},{"data":"2026-01-22","lote":2537,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 04 - FMT","linha_planilha":57},{"data":"2026-01-23","lote":2538,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 05 - FMT","linha_planilha":58},{"data":"2026-01-23","lote":2539,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 05 - FMT","linha_planilha":59},{"data":"2026-01-26","lote":2540,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 05 - FMT","linha_planilha":60},{"data":"2026-01-26","lote":2541,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 05 - FMT","linha_planilha":61},{"data":"2026-01-27","lote":2542,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 06 - FMT","linha_planilha":62},{"data":"2026-01-27","lote":2543,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 05 - FMT","linha_planilha":63},{"data":"2026-01-28","lote":2544,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 06 - FMT","linha_planilha":64},{"data":"2026-01-28","lote":2545,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 06 - FMT","linha_planilha":65},{"data":"2026-01-29","lote":2546,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 06 - FMT","linha_planilha":66},{"data":"2026-01-29","lote":2547,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 06 - FMT","linha_planilha":67},{"data":"2026-01-30","lote":2548,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 07 - FMT","linha_planilha":68},{"data":"2026-02-04","lote":2549,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 03 - FN","linha_planilha":69},{"data":"2026-02-05","lote":2550,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 04 -FN","linha_planilha":70},{"data":"2026-02-06","lote":2551,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 04 -FN","linha_planilha":71},{"data":"2026-02-10","lote":2552,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 04 -FN","linha_planilha":72},{"data":"2026-02-11","lote":2553,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 04 -FN","linha_planilha":73},{"data":"2026-02-12","lote":2554,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 04 -FN","linha_planilha":74},{"data":"2026-02-13","lote":2555,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 04 -FN","linha_planilha":75},{"data":"2026-02-18","lote":2556,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":130,"serie_planilha":"Série 04 -FN","linha_planilha":76},{"data":"2026-02-18","lote":2556,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":145,"serie_planilha":"Série 04 -FN","linha_planilha":77},{"data":"2026-02-19","lote":2557,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 04 -FN","linha_planilha":78},{"data":"2026-02-20","lote":2558,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":84,"serie_planilha":"Série 04 -FN","linha_planilha":79},{"data":"2026-02-20","lote":2558,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":191,"serie_planilha":"Série 05 -FN","linha_planilha":80},{"data":"2026-02-23","lote":2559,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 05 -FN","linha_planilha":81},{"data":"2026-03-09","lote":2560,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":170,"serie_planilha":"Série 05 -FN","linha_planilha":82},{"data":"2026-03-10","lote":2562,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":170,"serie_planilha":"Série 05 -FN","linha_planilha":83},{"data":"2026-01-30","lote":2607,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 06 - FMT","linha_planilha":84},{"data":"2026-01-31","lote":2608,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 07 - FMT","linha_planilha":85},{"data":"2026-02-02","lote":2609,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 07 - FMT","linha_planilha":86},{"data":"2026-02-02","lote":2610,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 07 - FMT","linha_planilha":87},{"data":"2026-02-03","lote":2611,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 03 - MP","linha_planilha":88},{"data":"2026-02-03","lote":2612,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 07 - FMT","linha_planilha":89},{"data":"2026-02-04","lote":2613,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 03 - MP","linha_planilha":90},{"data":"2026-02-04","lote":2614,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 07 - FMT","linha_planilha":91},{"data":"2026-02-05","lote":2615,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 03 - MP","linha_planilha":92},{"data":"2026-02-05","lote":2616,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 08 - FMT","linha_planilha":93},{"data":"2026-02-06","lote":2617,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 03 - MP","linha_planilha":94},{"data":"2026-02-06","lote":2618,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 08 - FMT","linha_planilha":95},{"data":"2026-02-07","lote":2619,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 04 - MP","linha_planilha":96},{"data":"2026-02-07","lote":2620,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 08 - FMT","linha_planilha":97},{"data":"2026-02-09","lote":2621,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 08 - FMT","linha_planilha":98},{"data":"2026-02-09","lote":2622,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 05 - MP","linha_planilha":99},{"data":"2026-02-10","lote":2623,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 08 - FMT","linha_planilha":100},{"data":"2026-02-10","lote":2624,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 08 - FMT","linha_planilha":101},{"data":"2026-02-11","lote":2625,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 09 - FMT","linha_planilha":102},{"data":"2026-02-11","lote":2626,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 05 - MP","linha_planilha":103},{"data":"2026-02-12","lote":2627,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 09 - FMT","linha_planilha":104},{"data":"2026-02-12","lote":2628,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 05 - MP","linha_planilha":105},{"data":"2026-02-13","lote":2629,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 10 - FMT","linha_planilha":106},{"data":"2026-02-13","lote":2630,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 05 - MP","linha_planilha":107},{"data":"2026-02-16","lote":2631,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 05 - MP","linha_planilha":108},{"data":"2026-02-16","lote":2632,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 05 - MP","linha_planilha":109},{"data":"2026-02-17","lote":2633,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 05 - MP","linha_planilha":110},{"data":"2026-02-17","lote":2634,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 11 - FMT","linha_planilha":111},{"data":"2026-02-18","lote":2635,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 06 - MP","linha_planilha":112},{"data":"2026-02-18","lote":2636,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 12 - FMT","linha_planilha":113},{"data":"2026-02-19","lote":2637,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 06 - MP","linha_planilha":114},{"data":"2026-02-19","lote":2638,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 13 - FMT","linha_planilha":115},{"data":"2026-02-20","lote":2639,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 06 - MP","linha_planilha":116},{"data":"2026-02-20","lote":2640,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 06 - MP","linha_planilha":117},{"data":"2026-02-21","lote":2641,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 14 - FMT","linha_planilha":118},{"data":"2026-02-23","lote":2642,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 14 - FMT","linha_planilha":119},{"data":"2026-02-23","lote":2643,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 14 - FMT","linha_planilha":120},{"data":"2026-02-24","lote":2644,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 14 - FMT","linha_planilha":121},{"data":"2026-02-24","lote":2645,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 14 - FMT","linha_planilha":122},{"data":"2026-02-25","lote":2646,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 14 - FMT","linha_planilha":123},{"data":"2026-02-25","lote":2647,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 15 - FMT","linha_planilha":124},{"data":"2026-02-26","lote":2648,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 15 - FMT","linha_planilha":125},{"data":"2026-02-26","lote":2649,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 15 - FMT","linha_planilha":126},{"data":"2026-02-27","lote":2650,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 06 - MP","linha_planilha":127},{"data":"2026-02-27","lote":2651,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 15 - FMT","linha_planilha":128},{"data":"2026-03-02","lote":2652,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 06 - MP","linha_planilha":129},{"data":"2026-03-02","lote":2653,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 15 - FMT","linha_planilha":130},{"data":"2026-03-03","lote":2654,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 08 - MP","linha_planilha":131},{"data":"2026-03-03","lote":2655,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 15 - FMT","linha_planilha":132},{"data":"2026-03-04","lote":2656,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 12 - MP","linha_planilha":133},{"data":"2026-03-04","lote":2657,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 16 - FMT","linha_planilha":134},{"data":"2026-03-05","lote":2658,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 11 - MP","linha_planilha":135},{"data":"2026-03-05","lote":2659,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 16 - FMT","linha_planilha":136},{"data":"2026-03-06","lote":2660,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 10 - MP","linha_planilha":137},{"data":"2026-03-06","lote":2661,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 16 - FMT","linha_planilha":138},{"data":"2026-03-07","lote":2662,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 16 - FMT","linha_planilha":139},{"data":"2026-03-09","lote":2663,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 09 - MP","linha_planilha":140},{"data":"2026-03-09","lote":2664,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 16 - FMT","linha_planilha":141},{"data":"2026-03-10","lote":2665,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 07 - MP","linha_planilha":142},{"data":"2026-03-10","lote":2666,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 16 - FMT","linha_planilha":143},{"data":"2026-03-11","lote":2667,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":318,"serie_planilha":"Série 08 - MP","linha_planilha":144},{"data":"2026-03-11","lote":2668,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 17 - FMT","linha_planilha":145},{"data":"2026-03-12","lote":2669,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":318,"serie_planilha":"Série 08 - MP","linha_planilha":146},{"data":"2026-03-12","lote":2670,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 17 - FMT","linha_planilha":147},{"data":"2026-03-16","lote":2671,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":318,"serie_planilha":"Série 08 - MP","linha_planilha":148},{"data":"2026-03-13","lote":2672,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 17 - FMT","linha_planilha":149},{"data":"2026-03-14","lote":2673,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 17 - FMT","linha_planilha":150},{"data":"2026-03-16","lote":2674,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 17 - FMT","linha_planilha":151},{"data":"2026-03-17","lote":2675,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":312,"serie_planilha":"Série 08 - MP","linha_planilha":152},{"data":"2026-03-17","lote":2675,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T curvo","quantidade":11,"serie_planilha":"Série 08 - MP","linha_planilha":153},{"data":"2026-03-17","lote":2675,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T reto","quantidade":7,"serie_planilha":"Série 08 - MP","linha_planilha":154},{"data":"2026-03-17","lote":2676,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 17 - FMT","linha_planilha":155},{"data":"2026-03-18","lote":2678,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":324,"serie_planilha":"Série 22 - FMT","linha_planilha":156},{"data":"2026-03-11","lote":2862,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":170,"serie_planilha":"Série 05 -FN","linha_planilha":157},{"data":"2026-03-12","lote":2892,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":170,"serie_planilha":"Série 05 -FN","linha_planilha":158},{"data":"2026-03-13","lote":2893,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":170,"serie_planilha":"Série 06 - FN","linha_planilha":159},{"data":"2026-03-16","lote":2894,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":250,"serie_planilha":"Série 06 - FN","linha_planilha":160},{"data":"2026-03-17","lote":2895,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":260,"serie_planilha":"Série 06 - FN","linha_planilha":161},{"data":"2026-03-18","lote":2866,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 06 - FN","linha_planilha":162},{"data":"2026-03-18","lote":2677,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":300,"serie_planilha":"Série 08 - MP","linha_planilha":163},{"data":"2026-03-18","lote":2677,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T curvo","quantidade":11,"serie_planilha":"Série 08 - MP","linha_planilha":164},{"data":"2026-03-18","lote":2677,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T reto","quantidade":19,"serie_planilha":"Série 08 - MP","linha_planilha":165},{"data":"2026-03-19","lote":2680,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 23 - FMT","linha_planilha":166},{"data":"2026-03-19","lote":2868,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 07 - FN","linha_planilha":167},{"data":"2026-03-20","lote":2681,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 21 - FMT","linha_planilha":168},{"data":"2026-03-20","lote":2679,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":300,"serie_planilha":"Série 08 - MP","linha_planilha":169},{"data":"2026-03-20","lote":2679,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T reto TR-68","quantidade":11,"serie_planilha":"Série 08 - MP","linha_planilha":170},{"data":"2026-03-20","lote":2679,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T reto TR-68","quantidade":19,"serie_planilha":"Série 08 - MP","linha_planilha":171},{"data":"2026-03-20","lote":2870,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":114,"serie_planilha":"Série 07 - FN","linha_planilha":172},{"data":"2026-03-20","lote":2870,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":161,"serie_planilha":"Série 07 - FN","linha_planilha":173},{"data":"2026-03-21","lote":2736,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 20 - FMT","linha_planilha":174},{"data":"2026-03-23","lote":2737,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":300,"serie_planilha":"Série 09 - MP","linha_planilha":175},{"data":"2026-03-23","lote":2737,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":11,"serie_planilha":"Série 09 - MP","linha_planilha":176},{"data":"2026-03-23","lote":2737,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":19,"serie_planilha":"Série 09 - MP","linha_planilha":177},{"data":"2026-03-23","lote":2738,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 19 - FMT","linha_planilha":178},{"data":"2026-03-24","lote":2739,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":312,"serie_planilha":"Série 09 - MP","linha_planilha":179},{"data":"2026-03-24","lote":2739,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T reto TR-68","quantidade":18,"serie_planilha":"Série 09 - MP","linha_planilha":180},{"data":"2026-03-24","lote":2740,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 18 - FMT","linha_planilha":181},{"data":"2026-03-23","lote":2873,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":270,"serie_planilha":"Série 07 - FN","linha_planilha":182},{"data":"2026-03-24","lote":2875,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":270,"serie_planilha":"Série 07 - FN","linha_planilha":183},{"data":"2026-03-25","lote":2741,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":312,"serie_planilha":"Série 09 - MP","linha_planilha":184},{"data":"2026-03-25","lote":2741,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T reto TR-68","quantidade":8,"serie_planilha":"Série 09 - MP","linha_planilha":185},{"data":"2026-03-25","lote":2741,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T reto UIC-60","quantidade":10,"serie_planilha":"Série 09 - MP","linha_planilha":186},{"data":"2026-03-25","lote":2742,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 24 - FMT","linha_planilha":187},{"data":"2026-03-25","lote":2874,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":270,"serie_planilha":"Série 07 - FN","linha_planilha":188},{"data":"2026-03-26","lote":2746,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":318,"serie_planilha":"Série 09 - MP","linha_planilha":189},{"data":"2026-03-26","lote":2746,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP C.T reto UIC-60","quantidade":12,"serie_planilha":"Série 09 - MP","linha_planilha":190},{"data":"2026-03-26","lote":2744,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 24 - FMT","linha_planilha":191},{"data":"2026-03-26","lote":2876,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 08 - FN","linha_planilha":192},{"data":"2026-03-27","lote":2749,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":324,"serie_planilha":"Série 09 - MP","linha_planilha":193},{"data":"2026-03-27","lote":2747,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 24 - FMT","linha_planilha":194},{"data":"2026-03-27","lote":2877,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 09 - FN","linha_planilha":195},{"data":"2026-03-28","lote":2750,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 24 - FMT","linha_planilha":196},{"data":"2026-03-30","lote":2752,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 10 - MP","linha_planilha":197},{"data":"2026-03-30","lote":2754,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 24 - FMT","linha_planilha":198},{"data":"2026-03-30","lote":2878,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Lote travado - FN","linha_planilha":199},{"data":"2026-03-31","lote":2757,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 10 - MP","linha_planilha":200},{"data":"2026-03-31","lote":2755,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 24 - FMT","linha_planilha":201},{"data":"2026-03-31","lote":2879,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 10 - FN","linha_planilha":202},{"data":"2026-04-01","lote":2760,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 10 - MP","linha_planilha":203},{"data":"2026-04-01","lote":2758,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":318,"serie_planilha":"Série 25 - FMT","linha_planilha":204},{"data":"2026-04-01","lote":2880,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 10 - FN","linha_planilha":205},{"data":"2026-04-02","lote":2763,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 10 - MP","linha_planilha":206},{"data":"2026-04-02","lote":2761,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":318,"serie_planilha":"Série 25 - FMT","linha_planilha":207},{"data":"2026-04-02","lote":2881,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 10 - FN","linha_planilha":208},{"data":"2026-04-06","lote":2767,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 10 - MP","linha_planilha":209},{"data":"2026-04-06","lote":2765,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":318,"serie_planilha":"Série 25 - FMT","linha_planilha":210},{"data":"2026-04-06","lote":2882,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 10 - FN","linha_planilha":211},{"data":"2026-04-07","lote":2770,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 10 - MP","linha_planilha":212},{"data":"2026-04-07","lote":2768,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":318,"serie_planilha":"Série 25 - FMT","linha_planilha":213},{"data":"2026-04-07","lote":2883,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 10 - FN","linha_planilha":214},{"data":"2026-04-08","lote":2771,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 25 - FMT","linha_planilha":215},{"data":"2026-04-08","lote":2884,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":270,"serie_planilha":"Série 10 - FN","linha_planilha":216},{"data":"2026-04-09","lote":2773,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 11 - MP","linha_planilha":217},{"data":"2026-04-09","lote":2774,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 25 - FMT","linha_planilha":218},{"data":"2026-04-09","lote":2885,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":270,"serie_planilha":"Série 10 - FN","linha_planilha":219},{"data":"2026-04-10","lote":2778,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 11 - MP","linha_planilha":220},{"data":"2026-04-10","lote":2886,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 26 - FMT","linha_planilha":221},{"data":"2026-04-10","lote":2776,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 11 - FN","linha_planilha":222},{"data":"2026-04-13","lote":2782,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 11 - MP","linha_planilha":223},{"data":"2026-04-13","lote":2780,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 26 - FMT","linha_planilha":224},{"data":"2026-04-13","lote":2887,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 12 - FN","linha_planilha":225},{"data":"2026-04-14","lote":2785,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 11 - MP","linha_planilha":226},{"data":"2026-04-14","lote":2783,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 26 - FMT","linha_planilha":227},{"data":"2026-04-14","lote":2888,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 12 - FN","linha_planilha":228},{"data":"2026-04-15","lote":2788,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 11 - MP","linha_planilha":229},{"data":"2026-04-15","lote":2786,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 26 - FMT","linha_planilha":230},{"data":"2026-04-15","lote":2889,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":224,"serie_planilha":"Série 12 - FN","linha_planilha":231},{"data":"2026-04-15","lote":2889,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":51,"serie_planilha":"Série 12 - FN","linha_planilha":232},{"data":"2026-04-16","lote":2791,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 11 - MP","linha_planilha":233},{"data":"2026-04-16","lote":2789,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 27 - FMT","linha_planilha":234},{"data":"2026-04-16","lote":2890,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 12 - FN","linha_planilha":235},{"data":"2026-04-17","lote":2794,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"Série 12 - MP","linha_planilha":236},{"data":"2026-04-17","lote":2792,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":330,"serie_planilha":"Série 27 - FMT","linha_planilha":237},{"data":"2026-04-17","lote":2891,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 12 - FN","linha_planilha":238},{"data":"2026-04-18","lote":2795,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":50,"serie_planilha":"Série 27 - FMT","linha_planilha":239},{"data":"2026-04-18","lote":2795,"projeto":"FMT","tipo":"Bitola Larga FMT USP","quantidade":280,"serie_planilha":"Série 27 - FMT","linha_planilha":240},{"data":"2026-04-21","lote":2799,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"0","linha_planilha":241},{"data":"2026-04-21","lote":2797,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"0","linha_planilha":242},{"data":"2026-04-21","lote":2800,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 12 - FN","linha_planilha":243},{"data":"2026-04-22","lote":2803,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"0","linha_planilha":244},{"data":"2026-04-22","lote":2802,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":330,"serie_planilha":"0","linha_planilha":245},{"data":"2026-04-22","lote":2801,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"Série 12 - FN","linha_planilha":246},{"data":"2026-04-23","lote":2805,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":0,"serie_planilha":"0","linha_planilha":247},{"data":"2026-04-24","lote":2804,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":0,"serie_planilha":"0","linha_planilha":248},{"data":"2026-04-24","lote":2808,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":0,"serie_planilha":"0","linha_planilha":249},{"data":"2026-04-27","lote":2809,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":0,"serie_planilha":"0","linha_planilha":250},{"data":"2026-04-27","lote":2811,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":0,"serie_planilha":"0","linha_planilha":251},{"data":"2026-04-28","lote":2812,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":0,"serie_planilha":"0","linha_planilha":252},{"data":"2026-04-28","lote":2814,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":0,"serie_planilha":"0","linha_planilha":253},{"data":"2026-04-23","lote":2806,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"0","linha_planilha":254},{"data":"2026-04-24","lote":2807,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"0","linha_planilha":255},{"data":"2026-04-27","lote":2810,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"0","linha_planilha":256},{"data":"2026-04-28","lote":2813,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"0","linha_planilha":257},{"data":"2026-05-01","lote":2818,"projeto":"MALHA PAULISTA","tipo":"Bitola Larga MP","quantidade":324,"serie_planilha":"0","linha_planilha":258},{"data":"2026-05-01","lote":2819,"projeto":"FERRO NORTE","tipo":"Bitola Larga FN","quantidade":275,"serie_planilha":"0","linha_planilha":259}];

    const state = {
      rows: [],
      source: "dados embutidos",
      mode: "regra",
      releases: loadReleases()
    };

    const $ = (id) => document.getElementById(id);
    const nf = new Intl.NumberFormat("pt-BR");

    function clean(value) {
      return String(value ?? "").replace(/\s+/g, " ").trim();
    }

    function normalizeKey(value) {
      return clean(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
    }

    function parseQty(value) {
      if (typeof value === "number" && Number.isFinite(value)) return Math.round(value);
      const normalized = clean(value).replace(/\./g, "").replace(",", ".");
      const parsed = Number(normalized);
      return Number.isFinite(parsed) ? Math.round(parsed) : 0;
    }

    function excelDateToISO(value) {
      if (!value) return "";
      if (value instanceof Date) return value.toISOString().slice(0, 10);
      if (typeof value === "number") {
        const utc = Math.round((value - 25569) * 86400 * 1000);
        return new Date(utc).toISOString().slice(0, 10);
      }
      const text = clean(value);
      if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
      const br = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
      if (br) {
        const y = br[3].length === 2 ? `20${br[3]}` : br[3];
        return `${y}-${br[2].padStart(2, "0")}-${br[1].padStart(2, "0")}`;
      }
      return text;
    }

    function formatDate(iso) {
      if (!iso) return "-";
      if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso;
      return new Date(`${iso}T00:00:00`).toLocaleDateString("pt-BR");
    }

    function projectCode(project) {
      const key = normalizeKey(project);
      if (key.includes("FERRO")) return "FN";
      if (key.includes("MALHA")) return "MP";
      if (key.includes("FMT")) return "FMT";
      return key.split(" ").map(part => part[0]).join("").slice(0, 4) || "PRJ";
    }

    function normalizeSerieName(value, project) {
      const raw = clean(value);
      if (!raw || raw === "0") return `Série aberta / sem série - ${projectCode(project)}`;
      return raw
        .replace(/\s*-\s*/g, " - ")
        .replace(/Série\s+(\d+)/i, (_, n) => `Série ${String(Number(n)).padStart(2, "0")}`)
        .replace(/\s+/g, " ")
        .trim();
    }

    function safeKey(value) {
      return normalizeKey(value).replace(/[^A-Z0-9]+/g, "_");
    }

    function releaseKey(project, serieName, mode) {
      return `${mode}:${safeKey(project)}:${safeKey(serieName)}`;
    }

    function loadReleases() {
      try { return JSON.parse(localStorage.getItem("cavan_series_liberadas_v1") || "{}"); }
      catch { return {}; }
    }

    function saveReleases() {
      localStorage.setItem("cavan_series_liberadas_v1", JSON.stringify(state.releases));
    }

    function mapWorkbookRows(rawRows) {
      if (!rawRows || !rawRows.length) return [];
      const header = rawRows[0].map(normalizeKey);
      const findIndex = (aliases) => header.findIndex(h => aliases.some(alias => h.includes(alias)));
      const idx = {
        data: findIndex(["DATA DE FABRICACAO", "DATA"]),
        lote: findIndex(["LOTE"]),
        projeto: findIndex(["PROJETO"]),
        tipo: findIndex(["TIPO DE DORMENTE", "TIPO"]),
        quantidade: findIndex(["TOTAL DA PRODUCAO", "TOTAL", "QUANTIDADE"]),
        serie: findIndex(["SERIE", "ENSAIO"])
      };
      return rawRows.slice(1).map((row, i) => {
        const projeto = clean(row[idx.projeto]);
        const lote = parseQty(row[idx.lote]);
        if (!projeto || !lote) return null;
        return {
          data: excelDateToISO(row[idx.data]),
          lote,
          projeto,
          tipo: clean(row[idx.tipo]),
          quantidade: parseQty(row[idx.quantidade]),
          serie_planilha: clean(row[idx.serie]),
          linha_planilha: i + 2
        };
      }).filter(Boolean);
    }

    function rowUid(row, idx) {
      return `${row.projeto}-${row.lote}-${row.data}-${idx}`;
    }

    function prepareRows(rows) {
      return rows.map((row, idx) => ({
        ...row,
        _uid: rowUid(row, idx),
        quantidade: parseQty(row.quantidade),
        lote: parseQty(row.lote),
        projeto: clean(row.projeto),
        tipo: clean(row.tipo),
        serie_planilha: clean(row.serie_planilha),
        data: excelDateToISO(row.data),
        _idx: idx
      })).filter(row => row.projeto && row.lote);
    }

    function sortRows(rows) {
      return [...rows].sort((a, b) => {
        const da = a.data || "9999-99-99";
        const db = b.data || "9999-99-99";
        if (da !== db) return da.localeCompare(db);
        if (a.lote !== b.lote) return a.lote - b.lote;
        return a._idx - b._idx;
      });
    }

    function emptySerie(project, name, mode) {
      return {
        project,
        name,
        mode,
        items: [],
        total: 0,
        productiveLots: 0,
        closed: false,
        plannedOnly: false,
        testLotUid: null
      };
    }

    function finalizeSerie(serie) {
      serie.total = serie.items.reduce((sum, row) => sum + Math.max(0, row.quantidade || 0), 0);
      serie.productiveLots = serie.items.filter(row => (row.quantidade || 0) > 0).length;
      if (!serie.plannedOnly) {
        serie.closed = serie.total >= LIMIT_QTY || serie.productiveLots >= LIMIT_LOTS;
        if (serie.closed && !serie.testLotUid) {
          const lastPositive = [...serie.items].reverse().find(row => (row.quantidade || 0) > 0);
          serie.testLotUid = lastPositive?._uid || null;
        }
      }
      serie.key = releaseKey(serie.project, serie.name, serie.mode);
      serie.released = Boolean(state.releases[serie.key]);
      const status = getSerieStatus(serie);
      serie.status = status.value;
      serie.statusLabel = status.label;
      return serie;
    }

    function buildByRule(rows) {
      const groups = groupBy(rows, row => row.projeto);
      return Object.entries(groups).map(([project, projectRows]) => {
        const sorted = sortRows(projectRows);
        const series = [];
        const planned = [];
        let seq = 1;
        let current = emptySerie(project, `Série ${String(seq).padStart(2, "0")} - ${projectCode(project)}`, "regra");

        sorted.forEach(row => {
          if ((row.quantidade || 0) <= 0) {
            planned.push(row);
            return;
          }
          current.items.push(row);
          current.total += row.quantidade;
          current.productiveLots += 1;
          if (current.total >= LIMIT_QTY || current.productiveLots >= LIMIT_LOTS) {
            current.testLotUid = row._uid;
            series.push(finalizeSerie(current));
            seq += 1;
            current = emptySerie(project, `Série ${String(seq).padStart(2, "0")} - ${projectCode(project)}`, "regra");
          }
        });

        if (current.items.length) series.push(finalizeSerie(current));
        if (planned.length) {
          const plannedSerie = emptySerie(project, `Lotes planejados / sem produção - ${projectCode(project)}`, "regra");
          plannedSerie.items = sortRows(planned);
          plannedSerie.plannedOnly = true;
          series.push(finalizeSerie(plannedSerie));
        }
        return finalizeProject(project, series);
      });
    }

    function buildBySpreadsheet(rows) {
      const projectGroups = groupBy(rows, row => row.projeto);
      return Object.entries(projectGroups).map(([project, projectRows]) => {
        const seriesGroups = groupBy(projectRows, row => normalizeSerieName(row.serie_planilha, project));
        const series = Object.entries(seriesGroups).map(([name, items]) => {
          const serie = emptySerie(project, name, "planilha");
          serie.items = sortRows(items);
          if (name.includes("aberta / sem série") && serie.items.every(row => (row.quantidade || 0) <= 0)) serie.plannedOnly = true;
          return finalizeSerie(serie);
        }).sort((a, b) => naturalCompare(a.name, b.name));
        return finalizeProject(project, series);
      });
    }

    function finalizeProject(project, series) {
      const total = series.reduce((sum, serie) => sum + serie.total, 0);
      const lots = series.reduce((sum, serie) => sum + serie.productiveLots, 0);
      const needsTest = series.filter(s => s.status === "ensaio").length;
      const released = series.filter(s => s.status === "liberada").length;
      const near = series.filter(s => s.status === "proximo").length;
      return { project, series, total, lots, needsTest, released, near };
    }

    function getSerieStatus(serie) {
      if (serie.plannedOnly) return { value: "planejado", label: "Planejado / sem produção" };
      if (serie.released) return { value: "liberada", label: "Liberada p/ carregamento" };
      if (serie.closed) return { value: "ensaio", label: "Ensaio obrigatório" };
      if (serie.mode === "planilha" && !serie.name.includes("aberta / sem série") && serie.total < NEAR_QTY && serie.productiveLots < NEAR_LOTS) {
        return { value: "conferir", label: "Conferir série" };
      }
      if (serie.total >= NEAR_QTY || serie.productiveLots >= NEAR_LOTS) return { value: "proximo", label: "Próximo do ensaio" };
      return { value: "andamento", label: "Em andamento" };
    }

    function groupBy(items, getter) {
      return items.reduce((acc, item) => {
        const key = getter(item);
        acc[key] = acc[key] || [];
        acc[key].push(item);
        return acc;
      }, {});
    }

    function naturalCompare(a, b) {
      return a.localeCompare(b, "pt-BR", { numeric: true, sensitivity: "base" });
    }

    function pct(value, limit) {
      return Math.max(0, Math.min(100, (value / limit) * 100));
    }

    function escapeHtml(value) {
      return clean(value).replace(/[&<>"]/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch]));
    }

    function buildProjects() {
      const prepared = prepareRows(state.rows);
      const projects = state.mode === "planilha" ? buildBySpreadsheet(prepared) : buildByRule(prepared);
      return projects.sort((a, b) => naturalCompare(a.project, b.project));
    }

    function render() {
      state.mode = $("modeSelect").value;
      const projects = buildProjects();
      fillProjectFilter(projects);

      const selectedProject = $("projectFilter").value;
      const selectedStatus = $("statusFilter").value;
      const term = normalizeKey($("searchInput").value);

      const filteredProjects = projects.map(project => {
        let series = project.series;
        if (selectedStatus !== "todos") series = series.filter(serie => serie.status === selectedStatus);
        if (term) {
          series = series.map(serie => ({
            ...serie,
            items: serie.items.filter(row => normalizeKey(`${row.lote} ${row.projeto} ${row.tipo} ${row.serie_planilha}`).includes(term))
          })).filter(serie => serie.items.length || normalizeKey(`${serie.name} ${serie.project}`).includes(term));
        }
        return { ...project, series };
      }).filter(project => (selectedProject === "todos" || project.project === selectedProject) && project.series.length);

      renderSummary(projects);
      renderProjects(filteredProjects);
      $("sourceNote").innerHTML = `Fonte: <strong>${escapeHtml(state.source)}</strong> • ${nf.format(prepareRows(state.rows).length)} linhas válidas lidas. Modo atual: <strong>${state.mode === "regra" ? "recalcular pela regra 2.000/10" : "usar coluna de série da planilha"}</strong>.`;
    }

    function fillProjectFilter(projects) {
      const select = $("projectFilter");
      const current = select.value || "todos";
      const options = ["todos", ...projects.map(p => p.project)];
      const existing = [...select.options].map(opt => opt.value).join("|");
      if (existing === options.join("|")) return;
      select.innerHTML = options.map(value => `<option value="${escapeHtml(value)}">${value === "todos" ? "Todos" : escapeHtml(value)}</option>`).join("");
      select.value = options.includes(current) ? current : "todos";
    }

    function renderSummary(projects) {
      const allSeries = projects.flatMap(p => p.series);
      const total = projects.reduce((sum, p) => sum + p.total, 0);
      const lots = projects.reduce((sum, p) => sum + p.lots, 0);
      const needsTest = allSeries.filter(s => s.status === "ensaio").length;
      const near = allSeries.filter(s => s.status === "proximo").length;
      const released = allSeries.filter(s => s.status === "liberada").length;
      const cards = [
        ["Projetos", projects.length],
        ["Dormentes produzidos", nf.format(total)],
        ["Lotes produzidos", nf.format(lots)],
        ["Séries c/ ensaio obrigatório", nf.format(needsTest)],
        ["Séries liberadas", nf.format(released)]
      ];
      $("summaryGrid").innerHTML = cards.map(([label, value]) => `<article class="kpi"><span>${label}</span><strong>${value}</strong></article>`).join("");
    }

    function renderProjects(projects) {
      if (!projects.length) {
        $("projects").innerHTML = `<div class="empty">Nenhum projeto encontrado com os filtros aplicados.</div>`;
        return;
      }
      $("projects").innerHTML = projects.map(project => renderProjectCard(project)).join("");
    }

    function renderProjectCard(project) {
      const statusText = project.needsTest
        ? `${project.needsTest} série(s) aguardando ensaio`
        : project.near
          ? `${project.near} série(s) próxima(s) do ensaio`
          : "Sem gatilho imediato";
      return `<article class="project-card">
        <header class="project-head">
          <div>
            <h2 class="project-title">${escapeHtml(project.project)}</h2>
            <p class="project-meta">${escapeHtml(statusText)}. Após ensaio aprovado, marque a série como liberada para carregamento.</p>
          </div>
          <div class="project-stats">
            <span class="pill dark">${nf.format(project.total)} dormentes</span>
            <span class="pill">${nf.format(project.lots)} lotes produzidos</span>
            <span class="pill">${nf.format(project.series.length)} séries/cards</span>
          </div>
        </header>
        <div class="series-grid">${project.series.map(renderSerieCard).join("")}</div>
      </article>`;
    }

    function renderSerieCard(serie) {
      const qtyPct = pct(serie.total, LIMIT_QTY);
      const lotPct = pct(serie.productiveLots, LIMIT_LOTS);
      const lastPositive = [...serie.items].reverse().find(row => (row.quantidade || 0) > 0);
      const candidate = serie.testLotUid
        ? serie.items.find(row => row._uid === serie.testLotUid)
        : lastPositive;
      const candidateLabel = serie.status === "ensaio" || serie.status === "liberada" ? "Lote separado para ensaio" : "Último lote produzido / candidato";
      const canRelease = serie.closed && serie.status !== "liberada";
      const canUndo = serie.status === "liberada";
      return `<section class="serie-card ${serie.status}">
        <div class="serie-top">
          <div>
            <h3 class="serie-name">${escapeHtml(serie.name)}</h3>
            <div class="serie-sub">${nf.format(serie.productiveLots)} lote(s) produzidos • ${nf.format(serie.total)} dormentes</div>
          </div>
          <span class="badge ${serie.status}">${escapeHtml(serie.statusLabel)}</span>
        </div>
        <div class="progress-wrap">
          <div class="progress-row">
            <div class="progress-label"><span>Dormentes</span><span>${nf.format(serie.total)} / ${nf.format(LIMIT_QTY)}</span></div>
            <div class="bar"><span style="width:${qtyPct}%"></span></div>
          </div>
          <div class="progress-row">
            <div class="progress-label"><span>Lotes</span><span>${nf.format(serie.productiveLots)} / ${nf.format(LIMIT_LOTS)}</span></div>
            <div class="bar"><span style="width:${lotPct}%"></span></div>
          </div>
        </div>
        ${candidate ? `<div class="candidate"><strong>${candidateLabel}:</strong> lote ${escapeHtml(candidate.lote)} • ${formatDate(candidate.data)} • ${nf.format(candidate.quantidade)} dormentes.</div>` : ""}
        <div class="serie-actions">
          ${canRelease ? `<button class="mini-btn release" data-action="release" data-key="${escapeHtml(serie.key)}">Marcar ensaio feito e liberar</button>` : ""}
          ${canUndo ? `<button class="mini-btn" data-action="undo" data-key="${escapeHtml(serie.key)}">Desfazer liberação</button>` : ""}
          ${serie.status === "conferir" ? `<span class="pill">Série abaixo do gatilho: validar na planilha</span>` : ""}
        </div>
        <details ${serie.status === "ensaio" || serie.status === "proximo" ? "open" : ""}>
          <summary>Ver lotes da série</summary>
          <table class="lot-table">
            <thead><tr><th>Data</th><th>Lote</th><th>Tipo</th><th>Qtd.</th><th>Série planilha</th></tr></thead>
            <tbody>${serie.items.map(row => renderLotRow(row, serie)).join("")}</tbody>
          </table>
        </details>
      </section>`;
    }

    function renderLotRow(row, serie) {
      const isTest = row._uid === serie.testLotUid;
      const isZero = (row.quantidade || 0) <= 0;
      return `<tr class="${isTest ? "test-lot" : ""} ${isZero ? "zero-lot" : ""}">
        <td>${formatDate(row.data)}</td>
        <td>${escapeHtml(row.lote)}${isTest ? `<span class="lot-badge">Ensaio</span>` : ""}</td>
        <td>${escapeHtml(row.tipo)}</td>
        <td>${nf.format(row.quantidade || 0)}</td>
        <td>${escapeHtml(row.serie_planilha || "-")}</td>
      </tr>`;
    }

    async function parseWorkbookFile(fileOrBuffer, label) {
      if (!window.XLSX) throw new Error("Biblioteca XLSX não carregada.");
      const buffer = fileOrBuffer instanceof ArrayBuffer ? fileOrBuffer : await fileOrBuffer.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
      const firstSheet = workbook.SheetNames[0];
      const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1, raw: true, defval: "" });
      const mapped = mapWorkbookRows(matrix);
      if (!mapped.length) throw new Error("Não encontrei linhas válidas na planilha.");
      state.rows = mapped;
      state.source = label;
      render();
    }

    async function tryLoadRepositorySpreadsheet() {
      state.rows = EMBEDDED_LOTES;
      state.source = "dados embutidos da planilha enviada";
      render();
      if (!window.XLSX) return;
      try {
        const response = await fetch(PLANILHA_URL, { cache: "no-store" });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const buffer = await response.arrayBuffer();
        await parseWorkbookFile(buffer, "Painel de series de DM.xlsx no repositório");
      } catch (error) {
        console.info("Usando dados embutidos; não foi possível carregar a planilha do repositório.", error);
      }
    }

    $("printBtn").addEventListener("click", () => window.print());
    $("fileInput").addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try { await parseWorkbookFile(file, file.name); }
      catch (error) { alert(`Não consegui ler a planilha: ${error.message}`); }
      finally { event.target.value = ""; }
    });
    ["searchInput", "projectFilter", "statusFilter", "modeSelect"].forEach(id => $(id).addEventListener("input", render));
    $("resetReleasesBtn").addEventListener("click", () => {
      if (!confirm("Limpar todas as marcações de séries liberadas neste navegador?")) return;
      state.releases = {};
      saveReleases();
      render();
    });
    document.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action]");
      if (!button) return;
      const key = button.dataset.key;
      if (button.dataset.action === "release") state.releases[key] = true;
      if (button.dataset.action === "undo") delete state.releases[key];
      saveReleases();
      render();
    });

    tryLoadRepositorySpreadsheet();
  