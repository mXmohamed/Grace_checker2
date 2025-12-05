// Application GRACE - Analyseur d'erreurs multi-périodes
class GraceAnalyzer {
    constructor() {
        this.zipFiles = [];
        this.referenceData = null;
        this.allData = [];
        this.pmData = {};
        this.errorFamilies = {};
        this.selectedPMs = [];
        this.selectedFamilies = [];
        this.csvHeaders = null; // Stocker les en-têtes réels des colonnes A-O
        
        this.initializeEventListeners();
        this.initializeErrorFamilies();
    }

    initializeEventListeners() {
        // Gestion des fichiers ZIP
        document.getElementById('zipFiles').addEventListener('change', (e) => {
            this.zipFiles = Array.from(e.target.files);
            this.checkReadyState();
        });

        // Gestion du fichier de référence
        document.getElementById('referenceFile').addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.loadReferenceFile(e.target.files[0]);
            }
        });

        // Boutons
        document.getElementById('analyzeBtn').addEventListener('click', () => this.analyzeFiles());
        document.getElementById('resetBtn').addEventListener('click', () => this.reset());
        document.getElementById('analyzeAndVisualize').addEventListener('click', () => this.analyzeAndVisualize());
        document.getElementById('generateReportBtn').addEventListener('click', () => this.generateExcelReport());

        // Drag & Drop pour les fichiers ZIP
        const zipInput = document.getElementById('zipFiles');
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            zipInput.addEventListener(eventName, this.preventDefaults, false);
        });

        zipInput.addEventListener('dragenter', () => zipInput.classList.add('drag-over'));
        zipInput.addEventListener('dragleave', () => zipInput.classList.remove('drag-over'));
        zipInput.addEventListener('drop', (e) => {
            zipInput.classList.remove('drag-over');
            const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.zip'));
            if (files.length > 0) {
                this.zipFiles = files;
                this.checkReadyState();
            }
        });
    }

    preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    async initializeErrorFamilies() {
        // Charger le mapping depuis FAMILLE.csv
        try {
            const response = await fetch('./FAMILLE.csv');
            const csvText = await response.text();
            
            Papa.parse(csvText, {
                delimiter: ';',
                header: true,
                skipEmptyLines: true,
                complete: (results) => {
                    this.errorCodeToFamily = {};
                    this.errorFamilies = {};
                    
                    results.data.forEach(row => {
                        const code = row.CODE_ERREUR?.trim();
                        const family = row.FAMILLE?.trim();
                        
                        if (code && family) {
                            this.errorCodeToFamily[code] = family;
                            
                            // Construire la liste des familles uniques
                            if (!this.errorFamilies[family]) {
                                this.errorFamilies[family] = [];
                            }
                        }
                    });
                    
                    // Ajouter AUTRES pour les codes non trouvés
                    this.errorFamilies['AUTRES'] = [];
                    
                    console.log(`Chargé ${Object.keys(this.errorCodeToFamily).length} codes d'erreur avec leurs familles`);
                }
            });
        } catch (error) {
            console.warn('Impossible de charger FAMILLE.csv, utilisation des familles par défaut');
            this.initializeDefaultFamilies();
        }
    }

    initializeDefaultFamilies() {
        // Mapping complet depuis FAMILLE.csv intégré directement dans l'application
        this.errorCodeToFamily = {
            // Mapping complet depuis FAMILLE.csv
            'ANO_CABLE_3': 'CABLE',
            'ANO_CABLE_4': 'CABLE',
            'ANO_CABLE_5': 'CABLE',
            'ANO_CABLE_6': 'CABLE',
            'ANO_CABLE_7': 'CABLE',
            'ANO_CABLE_8': 'CABLE',
            'ARMOIRE_122_A': 'ARMOIRE',
            'ARMOIRE_122_B': 'ARMOIRE',
            'ARMOIRE_37_A': 'ARMOIRE',
            'ARMOIRE_37_B': 'ARMOIRE',
            'ARMOIRE_37_C': 'ARMOIRE',
            'ARMOIRE_37_D': 'ARMOIRE',
            'ARMOIRE_37_E': 'ARMOIRE',
            'ARMOIRE_37_F': 'ARMOIRE',
            'ARMOIRE_38_A': 'ARMOIRE',
            'ARMOIRE_86_A': 'ARMOIRE',
            'ARMOIRE_86_B': 'ARMOIRE',
            'ARMOIRE_86_C': 'ARMOIRE',
            'ARMOIRE_86_E': 'ARMOIRE',
            'BAIE_120_A': 'BAIE',
            'BAIE_120_B': 'BAIE',
            'BAIE_140_A': 'BAIE',
            'BAIE_140_B': 'BAIE',
            'BAIE_52_A': 'BAIE',
            'BAIE_52_B': 'BAIE',
            'BAIE_90_A': 'BAIE',
            'BAIE_90_B': 'BAIE',
            'BPE_126_A': 'BPE',
            'BPE_126_B': 'BPE',
            'BPE_126_C': 'BPE',
            'BPE_126_D': 'BPE',
            'BPE_144_A': 'BPE',
            'BPE_24_A': 'BPE',
            'BPE_24_B': 'BPE',
            'BPE_25_A': 'BPE',
            'BPE_27_A': 'BPE',
            'BPE_28_A': 'BPE',
            'BPE_28_B': 'BPE',
            'BPE_30_A': 'BPE',
            'BPE_30_B': 'BPE',
            'BPE_30_C': 'BPE',
            'BPE_30_D': 'BPE',
            'BPE_31_A': 'BPE',
            'BPE_77_A': 'BPE',
            'BPE_77_B': 'BPE',
            'BPE_77_C': 'BPE',
            'BPE_78_A': 'BPE',
            'BPE_78_B': 'BPE',
            'BPE_78_C': 'BPE',
            'BPE_79_A': 'BPE',
            'BPE_80_A': 'BPE',
            'BPE_80_B': 'BPE',
            'BPE_80_C': 'BPE',
            'BPE_80_D': 'BPE',
            'BPE_80_E': 'BPE',
            'BPE_81_A': 'BPE',
            'BPE_96_A': 'BPE',
            'CABLE_10_A': 'CABLE',
            'CABLE_10_B': 'CABLE',
            'CABLE_119_A': 'CABLE',
            'CABLE_11_A': 'CABLE',
            'CABLE_11_B': 'CABLE',
            'CABLE_123_A': 'CABLE',
            'CABLE_123_B': 'CABLE',
            'CABLE_123_C': 'CABLE',
            'CABLE_123_D': 'CABLE',
            'CABLE_123_E': 'CABLE',
            'CABLE_123_F': 'CABLE',
            'CABLE_123_G': 'CABLE',
            'CABLE_123_H': 'CABLE',
            'CABLE_123_I': 'CABLE',
            'CABLE_123_J': 'CABLE',
            'CABLE_123_K': 'CABLE',
            'CABLE_123_L': 'CABLE',
            'CABLE_1_C': 'CABLE',
            'CABLE_1_D': 'CABLE',
            'CABLE_2_A': 'CABLE',
            'CABLE_2_B': 'CABLE',
            'CABLE_2_C': 'CABLE',
            'CABLE_3_A': 'CABLE',
            'CABLE_4_A': 'CABLE',
            'CABLE_4_B': 'CABLE',
            'CABLE_5_A': 'CABLE',
            'CABLE_5_B': 'CABLE',
            'CABLE_5_C': 'CABLE',
            'CABLE_5_D': 'CABLE',
            'CABLE_5_E': 'CABLE',
            'CABLE_5_F': 'CABLE',
            'CABLE_5_G': 'CABLE',
            'CABLE_5_H': 'CABLE',
            'CABLE_5_I': 'CABLE',
            'CABLE_5_J': 'CABLE',
            'CABLE_6_A': 'CABLE',
            'CABLE_6_B': 'CABLE',
            'CABLE_6_C': 'CABLE',
            'CABLE_6_D': 'CABLE',
            'CABLE_6_E': 'CABLE',
            'CABLE_6_F': 'CABLE',
            'CABLE_71_A': 'CABLE',
            'CABLE_71_B': 'CABLE',
            'CABLE_71_D': 'CABLE',
            'CABLE_71_E': 'CABLE',
            'CABLE_71_F': 'CABLE',
            'CABLE_71_G': 'CABLE',
            'CABLE_73_A': 'CABLE',
            'CABLE_74_A': 'CABLE',
            'CABLE_75_A': 'CABLE',
            'CABLE_7_A': 'CABLE',
            'CABLE_7_B': 'CABLE',
            'CABLE_7_C': 'CABLE',
            'CABLE_7_D': 'CABLE',
            'CABLE_7_E': 'CABLE',
            'CABLE_8_A': 'CABLE',
            'CABLE_8_B': 'CABLE',
            'CABLE_9_A': 'CABLE',
            'CASSETTE_118_A': 'CASSETTE',
            'CASSETTE_127_A': 'CASSETTE',
            'CASSETTE_142_A': 'CASSETTE',
            'CASSETTE_142_B': 'CASSETTE',
            'CASSETTE_55_A': 'CASSETTE',
            'CASSETTE_55_C': 'CASSETTE',
            'CASSETTE_55_D': 'CASSETTE',
            'CASSETTE_55_E': 'CASSETTE',
            'CASSETTE_93_A': 'CASSETTE',
            'CASSETTE_94_A': 'CASSETTE',
            'CHAMBRE_40_A': 'CHAMBRE',
            'CHAMBRE_40_B': 'CHAMBRE',
            'CHAMBRE_40_C': 'CHAMBRE',
            'CHAMBRE_40_D': 'CHAMBRE',
            'CHAMBRE_40_F': 'CHAMBRE',
            'CHAMBRE_40_G': 'CHAMBRE',
            'CHAMBRE_41_A': 'CHAMBRE',
            'CHAMBRE_41_B': 'CHAMBRE',
            'CHAMBRE_41_C': 'CHAMBRE',
            'CHAMBRE_41_D': 'CHAMBRE',
            'CHAMBRE_41_E': 'CHAMBRE',
            'CHAMBRE_41_F': 'CHAMBRE',
            'CHAMBRE_41_G': 'CHAMBRE',
            'CHAMBRE_41_H': 'CHAMBRE',
            'CHAMBRE_41_I': 'CHAMBRE',
            'CHAMBRE_41_J': 'CHAMBRE',
            'CHAMBRE_41_K': 'CHAMBRE',
            'CHAMBRE_41_L': 'CHAMBRE',
            'CHAMBRE_42_A': 'CHAMBRE',
            'CHAMBRE_43_A': 'CHAMBRE',
            'CHAMBRE_44_A': 'CHAMBRE',
            'CHAMBRE_44_B': 'CHAMBRE',
            'CHAMBRE_44_C': 'CHAMBRE',
            'CHAMBRE_44_D': 'CHAMBRE',
            'CHAMBRE_45_A': 'CHAMBRE',
            'CHAMBRE_87_A': 'CHAMBRE',
            'CHAMBRE_87_B': 'CHAMBRE',
            'CHAMBRE_88_A': 'CHAMBRE',
            'CHAMBRE_88_B': 'CHAMBRE',
            'CHAMBRE_88_C': 'CHAMBRE',
            'CHAMBRE_88_D': 'CHAMBRE',
            'CHAMBRE_88_E': 'CHAMBRE',
            'IPE_111_A': 'IPE',
            'IPE_111_B': 'IPE',
            'IPE_111_C': 'IPE',
            'IPE_111_D': 'IPE',
            'IPE_111_E': 'IPE',
            'IPE_111_F': 'IPE',
            'IPE_111_G': 'IPE',
            'IPE_111_H': 'IPE',
            'IPE_111_I': 'IPE',
            'IPE_111_J': 'IPE',
            'IPE_112_A': 'IPE',
            'IPE_112_B': 'IPE',
            'IPE_112_C': 'IPE',
            'IPE_133_A': 'IPE',
            'IPE_133_B': 'IPE',
            'IPE_133_C': 'IPE',
            'IPE_139_A': 'IPE',
            'IPE_139_B': 'IPE',
            'IPE_139_C': 'IPE',
            'IPE_139_D': 'IPE',
            'IPE_139_E': 'IPE',
            'IPE_139_F': 'IPE',
            'JUMPER_114_A': 'JUMPER',
            'JUMPER_114_B': 'JUMPER',
            'JUMPER_114_C': 'JUMPER',
            'JUMPER_114_D': 'JUMPER',
            'MODULE_103_A': 'MODULE',
            'MODULE_103_B': 'MODULE',
            'MODULE_103_C': 'MODULE',
            'MODULE_104_A': 'MODULE',
            'MODULE_105_A': 'MODULE',
            'MODULE_141_A': 'JUMPER',
            'MODULE_63_A': 'MODULE',
            'MODULE_63_B': 'MODULE',
            'MODULE_63_C': 'MODULE',
            'MODULE_64_A': 'MODULE',
            'MODULE_64_B': 'JUMPER',
            'MODULE_65_A': 'MODULE',
            'MODULE_65_B': 'MODULE',
            'MODULE_66_A': 'MODULE',
            'MODULE_67_A': 'MODULE',
            'OPTCABLE_113_A': 'CABLE',
            'OPTCABLE_113_B': 'CABLE',
            'OPTCABLE_113_C': 'CABLE',
            'OPTCABLE_113_D': 'CABLE',
            'OPTCABLE_113_E': 'CABLE',
            'OPTCABLE_113_F': 'CABLE',
            'OPTCABLE_113_G': 'CABLE',
            'OPTRECEIVER_109_A': 'OPTRECEIVER',
            'OPTRECEIVER_110_A': 'OPTRECEIVER',
            'PBO_101_A': 'PBO',
            'PBO_102_A': 'PBO',
            'PBO_130_A': 'PBO',
            'PBO_130_B': 'PBO',
            'PBO_130_C': 'PBO',
            'PBO_134_A': 'PBO',
            'PBO_134_B': 'PBO',
            'PBO_134_C': 'PBO',
            'PBO_134_E': 'PBO',
            'PBO_136_A': 'PBO',
            'PBO_136_B': 'PBO',
            'PBO_136_C': 'PBO',
            'PBO_136_D': 'PBO',
            'PBO_137_A': 'PBO',
            'PBO_137_B': 'PBO',
            'PBO_137_C': 'PBO',
            'PBO_137_D': 'PBO',
            'PBO_143_A': 'PBO',
            'PBO_60_A': 'PBO',
            'PBO_60_D': 'PBO',
            'PBO_60_E': 'PBO',
            'PBO_61_A': 'PBO',
            'PBO_62_A': 'PBO',
            'PBO_62_B': 'PBO',
            'PIGTAIL_115_A': 'PIGTAIL',
            'PIGTAIL_115_B': 'PIGTAIL',
            'PIGTAIL_115_C': 'PIGTAIL',
            'PIGTAIL_115_D': 'PIGTAIL',
            'PORTEE_AERIENNE_20_A': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_20_B': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_20_C': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_20_D': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_21_A': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_B': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_D': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_E': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_F': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_G': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_H': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_I': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_J': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_K': 'PORTEE_AERIENNE',
            'PORTEE_AERIENNE_22_L': 'PORTEE_AERIENNE',
            'POTEAU_46_A': 'POTEAU',
            'POTEAU_46_B': 'POTEAU',
            'POTEAU_46_M': 'POTEAU',
            'POTEAU_47_A': 'POTEAU',
            'POTEAU_47_B': 'POTEAU',
            'POTEAU_47_C': 'POTEAU',
            'POTEAU_48_A': 'POTEAU',
            'POTEAU_48_B': 'POTEAU',
            'POTEAU_49_A': 'POTEAU',
            'POTEAU_49_B': 'POTEAU',
            'POTEAU_50_A': 'POTEAU',
            'POTEAU_51_A': 'POTEAU',
            'POTEAU_89_A': 'POTEAU',
            'POTEAU_89_B': 'POTEAU',
            'POTEAU_89_C': 'POTEAU',
            'POTEAU_89_D': 'POTEAU',
            'POTEAU_89_E': 'POTEAU',
            'PRISE_53_A': 'PRISE',
            'PRISE_53_B': 'PRISE',
            'PRISE_91_A': 'PRISE',
            'PRISE_91_B': 'PRISE',
            'PRISE_92_A': 'PRISE',
            'PRISE_92_B': 'PRISE',
            'ROOM_100_B': 'ROOM',
            'ROOM_100_C': 'ROOM',
            'ROOM_100_D': 'ROOM',
            'ROOM_129_A': 'ROOM',
            'ROOM_57_A': 'ROOM',
            'ROOM_57_C': 'ROOM',
            'ROOM_57_D': 'ROOM',
            'ROOM_57_E': 'ROOM',
            'ROOM_57_F': 'ROOM',
            'ROOM_57_G': 'ROOM',
            'ROOM_57_H': 'ROOM',
            'ROOM_58_A': 'ROOM',
            'ROOM_58_B': 'ROOM',
            'ROOM_59_A': 'ROOM',
            'ROOM_70_A': 'ROOM',
            'ROOM_97_A': 'ROOM',
            'ROOM_98_A': 'ROOM',
            'ROOM_99_A': 'ROOM',
            'SITE_121_A': 'SITE',
            'SITE_121_B': 'SITE',
            'SITE_121_C': 'SITE',
            'SITE_128_A': 'SITE',
            'SITE_33_A': 'SITE',
            'SITE_33_C': 'SITE',
            'SITE_33_D': 'SITE',
            'SITE_33_E': 'SITE',
            'SITE_34_A': 'SITE',
            'SITE_69_A': 'SITE',
            'SITE_83_A': 'SITE',
            'SITE_83_B': 'SITE',
            'SITE_83_C': 'SITE',
            'SITE_83_D': 'SITE',
            'SITE_83_E': 'SITE',
            'SITE_83_F': 'SITE',
            'SITE_83_G': 'SITE',
            'SITE_84_A': 'SITE',
            'SITE_84_B': 'SITE',
            'SITE_84_C': 'SITE',
            'SITE_84_D': 'SITE',
            'SITE_84_E': 'SITE',
            'SITE_85_A': 'SITE',
            'SITE_85_B': 'SITE',
            'TIROIR_106_A': 'TIROIR',
            'TIROIR_106_B': 'TIROIR',
            'TIROIR_107_A': 'TIROIR',
            'TIROIR_108_A': 'TIROIR',
            'TIROIR_108_B': 'TIROIR',
            'TIROIR_135_A': 'TIROIR',
            'TIROIR_135_B': 'TIROIR',
            'TIROIR_138_A': 'TIROIR',
            'TRANCHEE_12_A': 'TRANCHEE',
            'TRANCHEE_12_B': 'TRANCHEE',
            'TRANCHEE_12_C': 'TRANCHEE',
            'TRANCHEE_12_D': 'TRANCHEE',
            'TRANCHEE_12_E': 'TRANCHEE',
            'TRANCHEE_12_F': 'TRANCHEE',
            'TRANCHEE_12_G': 'TRANCHEE',
            'TRANCHEE_12_H': 'TRANCHEE',
            'TRANCHEE_12_J': 'TRANCHEE',
            'TRANCHEE_12_K': 'TRANCHEE',
            'TRANCHEE_13_A': 'TRANCHEE',
            'TRANCHEE_13_B': 'TRANCHEE',
            'TRANCHEE_13_C': 'TRANCHEE',
            'TRANCHEE_13_D': 'TRANCHEE',
            'TRANCHEE_13_E': 'TRANCHEE',
            'TRANCHEE_14_A': 'TRANCHEE',
            'TRANCHEE_15_A': 'TRANCHEE',
            'TRANCHEE_16_A': 'TRANCHEE',
            'TRANCHEE_16_B': 'TRANCHEE',
            'TRANCHEE_16_C': 'TRANCHEE',
            'TRANCHEE_16_D': 'TRANCHEE',
            'TRANCHEE_16_E': 'TRANCHEE',
            'TRANCHEE_16_F': 'TRANCHEE',
            'TRANCHEE_17_A': 'TRANCHEE',
            'TRANCHEE_17_B': 'TRANCHEE',
            'TRANCHEE_19_A': 'TRANCHEE',
            'TRANCHEE_76_A': 'TRANCHEE',
            'TRANCHEE_76_B': 'TRANCHEE',
            'TRANCHEE_76_C': 'TRANCHEE',
            'TRANCHEE_76_D': 'TRANCHEE',
            'TRANCHEE_76_E': 'TRANCHEE',
            'TRANCHEE_76_F': 'TRANCHEE',
            'TRANCHEE_76_G': 'TRANCHEE',
            'TRANCHEE_76_H': 'TRANCHEE',
            'TRANCHEE_76_I': 'TRANCHEE',
            'TRANCHEE_76_J': 'TRANCHEE',
            'TRANCHEE_76_K': 'TRANCHEE',
            'ZAPM_124_A': 'ZAPM',
            'ZAPM_124_B': 'ZAPM',
            'ZNRO_125_A': 'ZNRO',
            'ZNRO_128_B': 'ZNRO'
        };
        
        // Construire la liste des familles uniques
        this.errorFamilies = {};
        Object.values(this.errorCodeToFamily).forEach(family => {
            if (!this.errorFamilies[family]) {
                this.errorFamilies[family] = [];
            }
        });
        
        // Ajouter AUTRES pour les codes non trouvés
        this.errorFamilies['AUTRES'] = [];
        
        console.log(`${Object.keys(this.errorCodeToFamily).length} codes d'erreur intégrés avec leurs familles`);
        console.log('Familles disponibles:', Object.keys(this.errorFamilies));
        console.log('Exemple de mapping:', Object.entries(this.errorCodeToFamily).slice(0, 5));
    }

    checkReadyState() {
        const analyzeBtn = document.getElementById('analyzeBtn');
        if (this.zipFiles.length > 0) {
            analyzeBtn.disabled = false;
            this.updateStatus(`${this.zipFiles.length} fichier(s) ZIP prêt(s) à analyser`);
        } else {
            analyzeBtn.disabled = true;
        }
    }

    async loadReferenceFile(file) {
        try {
            if (file.name.endsWith('.csv')) {
                const text = await file.text();
                Papa.parse(text, {
                    complete: (results) => {
                        this.processReferenceData(results.data);
                    },
                    header: true
                });
            } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                const data = await file.arrayBuffer();
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(data);
                const worksheet = workbook.worksheets[0];
                const rows = [];
                worksheet.eachRow((row, rowNumber) => {
                    rows.push(row.values.slice(1)); // Enlever le premier élément vide
                });
                this.processReferenceData(rows);
            }
        } catch (error) {
            this.showError('Erreur lors du chargement du fichier de référence: ' + error.message);
        }
    }

    processReferenceData(data) {
        // Améliorer la classification des familles d'erreurs basée sur le fichier de référence
        if (data && data.length > 0) {
            this.referenceData = data;
            this.updateErrorFamiliesFromReference(data);
            this.updateStatus('Fichier de référence chargé avec succès');
        }
    }

    updateErrorFamiliesFromReference(data) {
        // Analyser les codes d'erreur pour améliorer la classification
        data.forEach(row => {
            const errorCode = row[0] ? row[0].toString().toUpperCase() : '';
            if (errorCode) {
                let classified = false;
                for (const [family, keywords] of Object.entries(this.errorFamilies)) {
                    if (family !== 'AUTRES') {
                        for (const keyword of keywords) {
                            if (errorCode.includes(keyword)) {
                                classified = true;
                                break;
                            }
                        }
                        if (classified) break;
                    }
                }
            }
        });
    }

    async analyzeFiles() {
        this.showProgress(true, 'Analyse des fichiers ZIP en cours...');
        this.allData = [];
        this.pmData = {};

        try {
            let processedFiles = 0;
            const totalFiles = this.zipFiles.length;

            for (const zipFile of this.zipFiles) {
                await this.processZipFile(zipFile);
                processedFiles++;
                this.updateProgress((processedFiles / totalFiles) * 100, 
                    `Traitement: ${processedFiles}/${totalFiles} fichiers`);
            }

            this.organizeDataByPM();
            this.showFilterSection();
            this.showProgress(false);
            this.updateStatus(`Analyse terminée: ${Object.keys(this.pmData).length} PM trouvés`);

        } catch (error) {
            this.showError('Erreur lors de l\'analyse: ' + error.message);
            this.showProgress(false);
        }
    }

    async processZipFile(zipFile) {
        const zip = new JSZip();
        const content = await zip.loadAsync(zipFile);
        
        // Chercher le fichier de synthèse pour obtenir le REF_PM et la date
        let pmInfo = null;
        for (const [filename, file] of Object.entries(content.files)) {
            if (filename.includes('RAPPORT_SYNTHESE_GV3_SRO')) {
                const text = await file.async('text');
                pmInfo = this.extractPMInfo(text, filename);
                console.log('PM Info extraite:', pmInfo, 'depuis fichier:', filename);
                break;
            }
        }

        // Si pas de DSP trouvé dans le fichier de synthèse, chercher dans tous les CSV
        if (!pmInfo?.dsp) {
            console.log('DSP non trouvé dans fichier synthèse, recherche dans les CSV...');
            pmInfo = pmInfo || { refPM: null, date: null, dsp: null };
            
            for (const [filename, file] of Object.entries(content.files)) {
                if (filename.toLowerCase().endsWith('.csv') && !file.dir && !pmInfo.dsp) {
                    const text = await file.async('text');
                    const dsp = this.extractDSPFromCSV(text, filename);
                    if (dsp) {
                        pmInfo.dsp = dsp;
                        console.log('DSP trouvé dans', filename, ':', dsp);
                        break;
                    }
                }
            }
        }

        // Traiter tous les fichiers CSV
        for (const [filename, file] of Object.entries(content.files)) {
            if (filename.toLowerCase().endsWith('.csv') && !file.dir) {
                const text = await file.async('text');
                await this.processCSVContent(text, filename, pmInfo, zipFile.name);
            }
        }
    }

    extractPMInfo(content, filename) {
        // Extraire le REF_PM, la date et le DSP du fichier de synthèse
        const info = {
            refPM: null,
            date: null,
            dsp: null
        };

        // Extraire depuis le nom du fichier - patterns plus flexibles
        let nameMatch = filename.match(/SRO-[A-Z]+-\d+/);
        if (!nameMatch) {
            // Essayer d'autres patterns
            nameMatch = filename.match(/([A-Z]{2,5}-[A-Z]+-\d+)/);
        }
        if (nameMatch) {
            info.refPM = nameMatch[0];
        }

        // Extraire la date depuis le nom du fichier - patterns plus flexibles
        let dateMatch = filename.match(/(\d{14})/);
        if (!dateMatch) {
            // Essayer d'autres formats de date
            dateMatch = filename.match(/(\d{8})/); // YYYYMMDD
        }
        
        if (dateMatch) {
            if (dateMatch[1].length === 14) {
                info.date = this.parseTimestamp(dateMatch[1]);
            } else if (dateMatch[1].length === 8) {
                // Format YYYYMMDD - ajouter l'heure par défaut
                info.date = this.parseTimestamp(dateMatch[1] + '120000');
            }
        }

        // Essayer d'extraire depuis le contenu aussi
        const lines = content.split('\n');
        console.log('Contenu du fichier de synthèse (premières 10 lignes):');
        lines.slice(0, 10).forEach((line, i) => console.log(`Ligne ${i}: "${line}"`));
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            
            // Chercher le PM
            if (line.includes('SRO-') && !info.refPM) {
                const match = line.match(/SRO-[A-Z]+-\d+/);
                if (match) info.refPM = match[0];
            }
            
            // Chercher des dates dans le contenu
            if (!info.date) {
                const dateInContentMatch = line.match(/(\d{14})/);
                if (dateInContentMatch) {
                    info.date = this.parseTimestamp(dateInContentMatch[1]);
                }
            }
            
            // Chercher le DSP - essayer différents patterns
            if (line.includes('DSP') && (line.includes('Ref PM') || line.includes('Date') || line.includes('PM'))) {
                console.log(`Ligne avec en-tête DSP trouvée (${i}): "${line}"`);
                
                // La ligne suivante devrait contenir les valeurs
                if (i + 1 < lines.length) {
                    const dataLine = lines[i + 1].trim();
                    console.log(`Ligne de données suivante (${i+1}): "${dataLine}"`);
                    
                    if (dataLine) {
                        // Essayer différents séparateurs
                        const separators = ['\t', '    ', '   ', '  ', ' ', ';', ','];
                        for (const sep of separators) {
                            const parts = dataLine.split(sep).filter(p => p.trim());
                            console.log(`Test séparateur "${sep}": [${parts.join('|')}]`);
                            
                            if (parts.length >= 1) {
                                const dspCandidate = parts[0].trim();
                                console.log(`Candidat DSP: "${dspCandidate}"`);
                                
                                // Vérifier que c'est bien un DSP (2-5 lettres majuscules)
                                if (dspCandidate.match(/^[A-Z]{2,5}$/)) {
                                    console.log(`DSP valide trouvé: "${dspCandidate}"`);
                                    info.dsp = dspCandidate;
                                    
                                    // Aussi extraire PM et date si pas encore trouvés
                                    if (!info.refPM && parts.length >= 2 && parts[1].includes('SRO-')) {
                                        info.refPM = parts[1].trim();
                                    }
                                    if (!info.date && parts.length >= 3 && parts[2]) {
                                        const dateStr = parts[2].replace(/[,\.].*/,'').trim(); // Enlever les millisecondes
                                        if (dateStr.match(/^\d{14}$/)) {
                                            info.date = this.parseTimestamp(dateStr);
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                        
                        // Si toujours pas trouvé, essayer de prendre directement le premier mot
                        if (!info.dsp && dataLine.length > 0) {
                            const firstWord = dataLine.split(/\s+/)[0].trim();
                            console.log(`Premier mot de la ligne: "${firstWord}"`);
                            if (firstWord.match(/^[A-Z]{2,5}$/)) {
                                console.log(`DSP trouvé comme premier mot: "${firstWord}"`);
                                info.dsp = firstWord;
                            }
                        }
                    }
                }
            }
        }

        return info;
    }

    extractDSPFromCSV(csvContent, filename) {
        console.log('Recherche DSP dans fichier:', filename);
        
        // Détecter le séparateur
        const separators = [';', ',', '\t', '|'];
        let parseResult = null;
        
        for (const separator of separators) {
            const result = Papa.parse(csvContent, {
                delimiter: separator,
                header: false,
                skipEmptyLines: true,
                preview: 5 // Juste les premières lignes pour l'en-tête
            });
            
            if (result.data.length > 0) {
                parseResult = result;
                break;
            }
        }

        if (!parseResult || parseResult.data.length < 2) {
            console.log('Impossible de parser le CSV pour extraire DSP');
            return null;
        }

        // Chercher la colonne "DSP" dans l'en-tête (première ligne)
        const headerRow = parseResult.data[0];
        let dspColumnIndex = -1;
        
        for (let i = 0; i < headerRow.length; i++) {
            const header = headerRow[i]?.toString().trim().toUpperCase();
            if (header === 'DSP') {
                dspColumnIndex = i;
                console.log('Colonne DSP trouvée à l\'index:', i);
                break;
            }
        }
        
        if (dspColumnIndex === -1) {
            console.log('Colonne DSP non trouvée dans l\'en-tête:', headerRow);
            return null;
        }
        
        // Récupérer la valeur DSP de la première ligne de données
        if (parseResult.data.length > 1) {
            const dataRow = parseResult.data[1];
            if (dataRow[dspColumnIndex]) {
                const dsp = dataRow[dspColumnIndex].toString().trim();
                console.log('DSP extrait:', dsp);
                
                // Vérifier que c'est bien un DSP valide (2-5 lettres majuscules)
                if (dsp.match(/^[A-Z]{2,5}$/)) {
                    return dsp;
                } else {
                    console.log('DSP invalide (pas le bon format):', dsp);
                }
            }
        }
        
        return null;
    }

    parseTimestamp(timestamp) {
        // Format: 20250410094325 -> 2025-04-10 09:43:25
        if (timestamp && timestamp.length >= 14) {
            const year = timestamp.substr(0, 4);
            const month = timestamp.substr(4, 2);
            const day = timestamp.substr(6, 2);
            const hour = timestamp.substr(8, 2);
            const minute = timestamp.substr(10, 2);
            const second = timestamp.substr(12, 2);
            return new Date(`${year}-${month}-${day}T${hour}:${minute}:${second}`);
        }
        return null;
    }

    async processCSVContent(text, filename, pmInfo, zipName) {
        // Détecter le séparateur
        const separators = [';', ',', '\t', '|'];
        let parseResult = null;
        
        for (const separator of separators) {
            const result = Papa.parse(text, {
                delimiter: separator,
                header: false,
                skipEmptyLines: true
            });
            
            if (result.data.length > 0 && result.data[0].length >= 15) {
                parseResult = result;
                break;
            }
        }

        if (!parseResult) return;

        // Capturer les en-têtes du premier fichier traité (première ligne)
        if (!this.csvHeaders && parseResult.data.length > 0 && parseResult.data[0].length >= 15) {
            this.csvHeaders = parseResult.data[0].slice(0, 15); // Garder les 15 premières colonnes
            console.log('En-têtes capturés:', this.csvHeaders);
        }

        // Traiter les données
        parseResult.data.forEach((row, index) => {
            if (index === 0 || row.length < 15) return; // Ignorer l'en-tête
            
            // Essayer d'extraire le PM depuis différentes colonnes
            let refPM = pmInfo?.refPM;
            if (!refPM) {
                // Chercher dans les colonnes communes pour les PM
                refPM = row[6] || row[5] || row[1] || 'UNKNOWN'; // Colonnes G, F, B
                
                // Nettoyer le format du PM
                if (refPM && typeof refPM === 'string') {
                    const pmMatch = refPM.match(/[A-Z]{2,5}-[A-Z]+-\d+/);
                    if (pmMatch) {
                        refPM = pmMatch[0];
                    }
                }
            }
            
            // Essayer d'extraire la date depuis différentes sources
            let date = pmInfo?.date;
            if (!date) {
                // Essayer plusieurs colonnes pour la date
                for (let i = 14; i < Math.min(row.length, 20); i++) {
                    if (row[i] && typeof row[i] === 'string' && row[i].match(/\d{8,14}/)) {
                        date = this.parseTimestamp(row[i]);
                        if (date) break;
                    }
                }
            }
            
            // Si pas de date trouvée, utiliser la date du nom de fichier ZIP
            if (!date) {
                const zipDateMatch = zipName.match(/(\d{8,14})/);
                if (zipDateMatch) {
                    const dateStr = zipDateMatch[1];
                    date = this.parseTimestamp(dateStr.length === 8 ? dateStr + '120000' : dateStr);
                }
            }
            
            // Essayer d'obtenir le DSP : d'abord depuis pmInfo, sinon depuis la première colonne du CSV
            let dsp = pmInfo?.dsp;
            if (!dsp && row[0] && typeof row[0] === 'string') {
                const dspCandidate = row[0].toString().trim();
                // Vérifier que c'est bien un DSP (2-4 lettres majuscules)
                if (dspCandidate.match(/^[A-Z]{2,4}$/)) {
                    dsp = dspCandidate;
                }
            }
            
            const record = {
                refPM: refPM || 'UNKNOWN',
                date: date || new Date(),
                dsp: dsp || 'UNKNOWN', // DSP depuis le fichier de synthèse ou première colonne
                errorCode: row[3] || '', // Colonne D
                errorFamily: this.classifyError(row[3] || ''),
                zipFile: zipName,
                csvFile: filename,
                fullRow: row.slice(0, 15) // Garder les 15 premières colonnes
            };
            
            this.allData.push(record);
        });
    }

    classifyError(errorCode) {
        // Debug des codes d'erreur
        console.log('Classification du code:', errorCode);
        
        // Utiliser d'abord le mapping exact du fichier FAMILLE.csv
        if (this.errorCodeToFamily && this.errorCodeToFamily[errorCode]) {
            console.log('Trouvé via mapping exact:', this.errorCodeToFamily[errorCode]);
            return this.errorCodeToFamily[errorCode];
        }
        
        // Fallback: classification par mots-clés
        const upperCode = errorCode.toUpperCase();
        console.log('Code en majuscules:', upperCode);
        
        for (const [family, keywords] of Object.entries(this.errorFamilies)) {
            if (family !== 'AUTRES' && keywords.length > 0) {
                for (const keyword of keywords) {
                    if (upperCode.includes(keyword)) {
                        console.log('Trouvé via mot-clé:', keyword, '→', family);
                        return family;
                    }
                }
            }
        }
        
        console.log('Aucune famille trouvée, classification AUTRES');
        return 'AUTRES';
    }

    organizeDataByPM() {
        // Organiser les données par PM et par période
        this.allData.forEach(record => {
            if (!this.pmData[record.refPM]) {
                this.pmData[record.refPM] = {};
            }
            
            const period = this.getPeriodKey(record.date);
            if (!this.pmData[record.refPM][period]) {
                this.pmData[record.refPM][period] = [];
            }
            
            this.pmData[record.refPM][period].push(record);
        });
    }

    getPeriodKey(date) {
        if (!date) return 'UNKNOWN';
        // Utiliser la date complète pour distinguer les périodes
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
    }

    showFilterSection() {
        // Afficher la section des filtres
        document.getElementById('filterSection').style.display = 'block';
        
        // Remplir la liste des PM
        const pmSelect = document.getElementById('pmFilter');
        pmSelect.innerHTML = '';
        
        Object.keys(this.pmData).sort().forEach(pm => {
            const option = document.createElement('option');
            option.value = pm;
            option.textContent = `${pm} (${Object.keys(this.pmData[pm]).length} périodes)`;
            pmSelect.appendChild(option);
        });

        // Remplir les checkboxes des familles
        const familyDiv = document.getElementById('familyFilters');
        familyDiv.innerHTML = '';
        
        Object.keys(this.errorFamilies).forEach(family => {
            const div = document.createElement('div');
            div.className = 'form-check';
            div.innerHTML = `
                <input class="form-check-input" type="checkbox" value="${family}" 
                       id="family_${family}" checked>
                <label class="form-check-label" for="family_${family}">
                    ${family}
                </label>
            `;
            familyDiv.appendChild(div);
        });

        document.getElementById('analyzeAndVisualize').disabled = false;
    }

    async analyzeAndVisualize() {
        // Récupérer les sélections
        this.selectedPMs = Array.from(document.getElementById('pmFilter').selectedOptions)
            .map(opt => opt.value);
        
        if (this.selectedPMs.length === 0) {
            this.showError('Veuillez sélectionner au moins un PM');
            return;
        }

        this.selectedFamilies = Array.from(document.querySelectorAll('#familyFilters input:checked'))
            .map(cb => cb.value);

        this.showProgress(true, 'Analyse et visualisation en cours...');

        try {
            // Activer le bouton d'export Excel
            document.getElementById('generateReportBtn').disabled = false;
            
            // Afficher les résultats directement
            this.updateProgress(100, 'Analyse terminée');
            this.showProgress(false);
            this.showResults();

        } catch (error) {
            this.showError('Erreur lors de l\'analyse: ' + error.message);
            this.showProgress(false);
        }
    }

    async generateExcelReport() {
        // Vérifier que l'analyse a été faite
        if (this.selectedPMs.length === 0) {
            this.showError('Veuillez d\'abord effectuer l\'analyse en cliquant sur "Analyser et visualiser"');
            return;
        }

        this.showProgress(true, 'Génération du rapport Excel...');

        try {
            const workbook = new ExcelJS.Workbook();
            
            // Feuille 1: Données brutes (colonnes A-O concaténées)
            this.updateProgress(25, 'Création de la feuille des données brutes...');
            await this.createRawDataSheet(workbook);
            
            // Feuille 2: Comparatif temporel
            this.updateProgress(50, 'Création de la feuille de comparaison temporelle...');
            await this.createComparisonSheet(workbook);
            
            // Feuille 3: Données pour graphiques
            this.updateProgress(75, 'Création de la feuille des graphiques...');
            await this.createChartsDataSheet(workbook);

            // Sauvegarder le fichier
            this.updateProgress(90, 'Génération du fichier Excel...');
            const buffer = await workbook.xlsx.writeBuffer();
            this.downloadFile(buffer, `GRACE_Analyse_${new Date().toISOString().slice(0,10)}.xlsx`);
            
            this.showProgress(false);
            this.updateStatus('Rapport Excel généré avec succès');

        } catch (error) {
            this.showError('Erreur lors de la génération du rapport: ' + error.message);
            this.showProgress(false);
        }
    }

    // Ancienne méthode conservée pour compatibilité (maintenant inutilisée)
    async generateReport() {
        return this.generateExcelReport();
    }

    async createRawDataSheet(workbook) {
        const worksheet = workbook.addWorksheet('Données brutes');
        
        // En-têtes des 15 colonnes A-O + colonnes supplémentaires
        // Utiliser les vrais en-têtes s'ils sont disponibles, sinon fallback vers A-O
        const csvHeaders = this.csvHeaders && this.csvHeaders.length >= 15 
            ? this.csvHeaders.slice(0, 15) 
            : ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];
        const headers = [...csvHeaders, 'PM', 'Famille', 'Date Export', 'Fichier ZIP'];
        const headerRow = worksheet.addRow(headers);
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        };
        headerRow.font.color = { argb: 'FFFFFFFF' };

        // Filtrer les données selon les sélections
        const filteredData = this.allData.filter(record => 
            this.selectedPMs.includes(record.refPM) && 
            this.selectedFamilies.includes(record.errorFamily)
        );

        // Ajouter les données
        filteredData.forEach(record => {
            const row = [];
            
            // Les 15 colonnes A-O
            for (let i = 0; i < 15; i++) {
                row.push(record.fullRow[i] || '');
            }
            
            // Colonnes supplémentaires
            row.push(record.refPM);
            row.push(record.errorFamily);
            row.push(record.date ? record.date.toLocaleDateString('fr-FR') : 'N/A');
            row.push(record.zipFile);
            
            worksheet.addRow(row);
        });

        // Auto-ajuster les colonnes
        worksheet.columns.forEach((column, index) => {
            column.width = index < 15 ? 12 : (index === 15 ? 16 : 14);
        });
    }

    async createComparisonSheet(workbook) {
        const worksheet = workbook.addWorksheet('Comparatif temporel');
        
        // Obtenir les données de comparaison pour les dates
        const comparisonData = this.buildComparisonData();
        
        // Formater les dates pour les en-têtes
        const oldestDateStr = comparisonData.periodInfo?.oldestDate ? 
            comparisonData.periodInfo.oldestDate.toLocaleDateString('fr-FR') : 'ancienne date';
        const newestDateStr = comparisonData.periodInfo?.newestDate ? 
            comparisonData.periodInfo.newestDate.toLocaleDateString('fr-FR') : 'récente date';
        
        // En-têtes
        const headers = ['DSP', 'PM', 'Famille d\'erreur', 'Code d\'erreur', 'Description d\'erreur', 
                        `Nb erreurs (${oldestDateStr})`, `Nb erreurs (${newestDateStr})`, 'Amélioration', '% d\'évolution'];
        const headerRow = worksheet.addRow(headers);
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF70AD47' }
        };
        headerRow.font.color = { argb: 'FFFFFFFF' };

        // Les données de comparaison ont déjà été obtenues plus haut
        
        // Ajouter les données
        comparisonData.forEach(row => {
            const evolutionText = row.evolution > 0 ? 'DÉGRADATION' : 
                                  row.evolution < 0 ? 'AMÉLIORATION' : 'STABLE';
            
            const excelRow = worksheet.addRow([
                row.dsp || 'N/A',
                row.pm,
                row.family,
                row.errorCode,
                row.description,
                row.oldCount,
                row.newCount,
                evolutionText,
                row.evolutionPercent
            ]);
            
            // Coloration selon l'évolution
            const evolutionCell = excelRow.getCell(8);
            const percentCell = excelRow.getCell(9);
            
            if (row.evolution > 0) {
                evolutionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF9999' } };
                percentCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF9999' } };
            } else if (row.evolution < 0) {
                evolutionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF99FF99' } };
                percentCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF99FF99' } };
            } else {
                evolutionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };
                percentCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };
            }
        });

        // Auto-ajuster les colonnes
        worksheet.columns.forEach(column => {
            column.width = 16;
        });
        worksheet.getColumn(5).width = 25; // Description plus large
    }

    async createChartsDataSheet(workbook) {
        const worksheet = workbook.addWorksheet('Graphiques');
        
        // Section 1: Erreurs par famille et par période
        worksheet.addRow(['Données pour graphiques - Évolution temporelle par famille']);
        
        // Organiser les données par période et famille
        const periodData = {};
        const filteredData = this.allData.filter(record => 
            this.selectedPMs.includes(record.refPM) && 
            this.selectedFamilies.includes(record.errorFamily)
        );

        filteredData.forEach(record => {
            const period = this.getPeriodKey(record.date);
            if (!periodData[period]) {
                periodData[period] = {};
            }
            if (!periodData[period][record.errorFamily]) {
                periodData[period][record.errorFamily] = 0;
            }
            periodData[period][record.errorFamily]++;
        });

        const periods = Object.keys(periodData).sort();
        
        // En-têtes du tableau
        const headers = ['Période', ...this.selectedFamilies, 'Total'];
        const headerRow = worksheet.addRow(headers);
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE67E22' }
        };
        headerRow.font.color = { argb: 'FFFFFFFF' };

        // Ajouter les données
        periods.forEach(period => {
            const row = [period];
            let total = 0;
            
            this.selectedFamilies.forEach(family => {
                const count = periodData[period][family] || 0;
                row.push(count);
                total += count;
            });
            
            row.push(total);
            worksheet.addRow(row);
        });

        // Section 2: Erreurs par PM
        worksheet.addRow([]);
        worksheet.addRow(['Données pour graphiques - Erreurs par PM']);
        
        const pmCounts = {};
        this.selectedPMs.forEach(pm => {
            pmCounts[pm] = this.allData.filter(record => record.refPM === pm).length;
        });

        const pmHeaderRow = worksheet.addRow(['PM', 'Nombre d\'erreurs']);
        pmHeaderRow.font = { bold: true };
        pmHeaderRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF3498DB' }
        };
        pmHeaderRow.font.color = { argb: 'FFFFFFFF' };

        Object.entries(pmCounts).forEach(([pm, count]) => {
            worksheet.addRow([pm, count]);
        });

        // Auto-ajuster les colonnes
        worksheet.columns.forEach(column => {
            column.width = 14;
        });
        worksheet.getColumn(1).width = 18;
        
        // Ajouter des graphiques Excel natifs pour chaque PM
        this.addChartsToExcel(worksheet);
    }

    addChartsToExcel(worksheet) {
        // Trouver la première ligne libre après les données existantes
        let currentRow = worksheet.rowCount + 5; // 5 lignes d'espacement après les données
        
        // Créer un graphique pour chaque PM sélectionné
        this.selectedPMs.forEach((pm, pmIndex) => {
            // Préparer les données pour ce PM spécifique
            const chartData = this.prepareChartDataForPM(pm);
            
            if (chartData.families.length === 0) return; // Pas de données pour ce PM
            
            // Ajouter le titre du graphique
            const titleRow = worksheet.getRow(currentRow);
            titleRow.getCell(1).value = `Comparaison erreurs - PM: ${pm}`;
            titleRow.getCell(1).font = { size: 14, bold: true };
            currentRow += 2;
            
            // Créer le tableau de données pour le graphique
            const dataStartRow = currentRow;
            
            // En-têtes : Famille + chaque période (date)
            const headerRow = worksheet.getRow(currentRow);
            headerRow.getCell(1).value = 'Famille';
            chartData.periods.forEach((period, index) => {
                headerRow.getCell(index + 2).value = this.formatDateFromPeriod(period);
            });
            headerRow.font = { bold: true };
            currentRow++;
            
            // Données : une ligne par famille
            chartData.families.forEach(family => {
                const row = worksheet.getRow(currentRow);
                row.getCell(1).value = family;
                chartData.periods.forEach((period, index) => {
                    row.getCell(index + 2).value = chartData.data[family][period] || 0;
                });
                currentRow++;
            });
            
            const dataEndRow = currentRow - 1;
            const dataEndCol = 1 + chartData.periods.length;
            
            // Créer le graphique Excel natif
            this.createExcelChart(worksheet, pm, dataStartRow, dataEndRow, dataEndCol, chartData.periods, currentRow + 2);
            
            // Passer à la position pour le prochain graphique
            currentRow += 24; // 20 lignes pour le graphique + 4 lignes d'espacement
        });
    }
    
    prepareChartDataForPM(pm) {
        // Filtrer les données pour ce PM spécifique
        const pmData = this.allData.filter(record => 
            record.refPM === pm && 
            this.selectedFamilies.includes(record.errorFamily)
        );
        
        // Obtenir toutes les périodes pour ce PM
        const periods = [...new Set(pmData.map(record => this.getPeriodKey(record.date)))].sort();
        
        // Organiser les données par famille et période
        const familyPeriodData = {};
        pmData.forEach(record => {
            const family = record.errorFamily;
            const period = this.getPeriodKey(record.date);
            
            if (!familyPeriodData[family]) {
                familyPeriodData[family] = {};
            }
            if (!familyPeriodData[family][period]) {
                familyPeriodData[family][period] = 0;
            }
            familyPeriodData[family][period]++;
        });
        
        // Filtrer les familles qui ont au moins 1 erreur dans une période
        const validFamilies = Object.keys(familyPeriodData).filter(family => {
            return periods.some(period => familyPeriodData[family][period] > 0);
        });
        
        // Ordonner par nombre total d'erreurs décroissant
        validFamilies.sort((a, b) => {
            const totalA = periods.reduce((sum, period) => sum + (familyPeriodData[a][period] || 0), 0);
            const totalB = periods.reduce((sum, period) => sum + (familyPeriodData[b][period] || 0), 0);
            return totalB - totalA;
        });
        
        return {
            families: validFamilies,
            periods: periods,
            data: familyPeriodData
        };
    }
    
    createExcelChart(worksheet, pm, dataStartRow, dataEndRow, dataEndCol, periods, chartStartRow) {
        try {
            // Créer un graphique en barres groupées
            const chart = worksheet.addChart({
                type: 'col',
                name: `Chart_${pm.replace(/[^a-zA-Z0-9]/g, '_')}`,
                title: {
                    name: `Comparaison des erreurs - ${pm}`,
                    size: 12
                }
            });
            
            // Définir les couleurs selon le nombre de périodes
            const colors = this.getExcelChartColors(periods.length);
            
            // Ajouter une série pour chaque période
            periods.forEach((period, index) => {
                const date = this.formatDateFromPeriod(period);
                chart.addSeries({
                    name: date,
                    categories: `'Graphiques'!$A$${dataStartRow + 1}:$A$${dataEndRow}`,
                    values: `'Graphiques'!$${String.fromCharCode(66 + index)}$${dataStartRow + 1}:$${String.fromCharCode(66 + index)}$${dataEndRow}`,
                    fill: colors[index]
                });
            });
            
            // Configurer les axes
            chart.setXAxis({
                name: 'Familles d\'erreurs',
                nameRotation: -45
            });
            
            chart.setYAxis({
                name: 'Nombre d\'erreurs',
                min: 0
            });
            
            // Positionner le graphique (colonnes A-J, hauteur 20 lignes)
            chart.setPosition('A', chartStartRow, 'J', chartStartRow + 20);
            
            // Ajouter le graphique à la feuille
            worksheet.addChart(chart);
            
        } catch (error) {
            console.warn('Erreur lors de la création du graphique Excel:', error);
            // Ajouter une note si le graphique ne peut pas être créé
            const errorRow = worksheet.getRow(chartStartRow);
            errorRow.getCell(1).value = `Graphique non disponible pour ${pm} (${error.message})`;
            errorRow.getCell(1).font = { color: { argb: 'FFFF0000' } };
        }
    }
    
    getExcelChartColors(periodCount) {
        // Couleurs Excel : Gris pour première, Bleu pour dernière, autres couleurs pour intermédiaires
        const colors = [
            '808080', // Gris pour ancienne
            '4BC0C0', // Vert pour intermédiaire
            'FFCE56', // Jaune pour intermédiaire  
            '9966FF', // Violet pour intermédiaire
            '36A2EB'  // Bleu pour nouvelle
        ];
        
        if (periodCount === 1) {
            return [colors[4]]; // Bleu si une seule période
        } else if (periodCount === 2) {
            return [colors[0], colors[4]]; // Gris + Bleu
        } else {
            // Pour 3+ périodes : Gris, couleurs intermédiaires, Bleu
            const result = [colors[0]]; // Première = Gris
            for (let i = 1; i < periodCount - 1; i++) {
                result.push(colors[Math.min(i, colors.length - 2)]);
            }
            result.push(colors[4]); // Dernière = Bleu
            return result;
        }
    }

    downloadFile(buffer, filename) {
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    showResults() {
        document.getElementById('resultsSection').style.display = 'block';
        
        // Remplir les données pour chaque onglet
        this.populateRawDataTab();
        this.populateComparisonTab();
        this.populateChartsTab();
        
        // Scroll vers les résultats
        document.getElementById('resultsSection').scrollIntoView({ behavior: 'smooth' });
    }

    populateRawDataTab() {
        const rawDataContent = document.getElementById('rawDataContent');
        
        // Filtrer les données selon les sélections
        const filteredData = this.allData.filter(record => 
            this.selectedPMs.includes(record.refPM) && 
            this.selectedFamilies.includes(record.errorFamily)
        );

        // Debug pour les en-têtes
        console.log('csvHeaders disponibles:', this.csvHeaders);
        
        // Créer le tableau avec les 15 colonnes (vrais en-têtes ou A-O)
        let csvHeaders = this.csvHeaders && this.csvHeaders.length >= 15 
            ? this.csvHeaders.slice(0, 15) 
            : ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];
        
        // Supprimer le BOM du premier en-tête si présent
        if (csvHeaders[0] && csvHeaders[0].startsWith('\ufeff')) {
            csvHeaders[0] = csvHeaders[0].substring(1);
        }
        
        const headers = csvHeaders;
        
        console.log('En-têtes utilisés:', headers);
        
        let tableHTML = `
            <div class="table-responsive">
                <table class="table table-striped table-sm">
                    <thead class="table-dark">
                        <tr>
                            ${headers.map(h => `<th>${h}</th>`).join('')}
                            <th>PM</th>
                            <th>Famille</th>
                            <th>Date Export</th>
                            <th>Fichier ZIP</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        filteredData.forEach(record => {
            const row = record.fullRow || [];
            tableHTML += '<tr>';
            
            // Les 15 colonnes A-O
            for (let i = 0; i < 15; i++) {
                tableHTML += `<td>${row[i] || ''}</td>`;
            }
            
            // Colonnes supplémentaires
            tableHTML += `
                <td>${record.refPM}</td>
                <td><span class="badge bg-secondary">${record.errorFamily}</span></td>
                <td>${record.date ? record.date.toLocaleDateString('fr-FR') : 'N/A'}</td>
                <td><small>${record.zipFile}</small></td>
            `;
            tableHTML += '</tr>';
        });

        tableHTML += `
                    </tbody>
                </table>
            </div>
            <div class="mt-2">
                <small class="text-muted">
                    <i class="bi bi-info-circle"></i> 
                    ${filteredData.length} enregistrements affichés
                </small>
            </div>
        `;

        rawDataContent.innerHTML = tableHTML;
    }

    populateComparisonTab() {
        const comparisonContent = document.getElementById('comparisonContent');
        
        // Créer le comparatif temporel
        const comparisonData = this.buildComparisonData();
        
        // Formater les dates pour les en-têtes
        const oldestDateStr = comparisonData.periodInfo?.oldestDate ? 
            comparisonData.periodInfo.oldestDate.toLocaleDateString('fr-FR') : 'ancienne date';
        const newestDateStr = comparisonData.periodInfo?.newestDate ? 
            comparisonData.periodInfo.newestDate.toLocaleDateString('fr-FR') : 'récente date';
        
        let tableHTML = `
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead class="table-primary">
                        <tr>
                            <th>DSP</th>
                            <th>PM</th>
                            <th>Famille d'erreur</th>
                            <th>Code d'erreur</th>
                            <th>Description d'erreur</th>
                            <th>Nb erreurs (${oldestDateStr})</th>
                            <th>Nb erreurs (${newestDateStr})</th>
                            <th>Amélioration</th>
                            <th>% d'évolution</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        comparisonData.forEach(row => {
            const evolutionIcon = this.getEvolutionIcon(row.evolution);
            const evolutionClass = row.evolution > 0 ? 'text-danger' : row.evolution < 0 ? 'text-success' : 'text-warning';
            
            tableHTML += `
                <tr>
                    <td>${row.dsp || 'N/A'}</td>
                    <td>${row.pm}</td>
                    <td><span class="badge bg-secondary">${row.family}</span></td>
                    <td><code>${row.errorCode}</code></td>
                    <td>${row.description}</td>
                    <td><span class="badge bg-light text-dark">${row.oldCount}</span></td>
                    <td><span class="badge bg-light text-dark">${row.newCount}</span></td>
                    <td class="${evolutionClass}">${evolutionIcon}</td>
                    <td class="${evolutionClass}"><strong>${row.evolutionPercent}%</strong></td>
                </tr>
            `;
        });

        tableHTML += `
                    </tbody>
                </table>
            </div>
        `;

        comparisonContent.innerHTML = tableHTML;
    }

    buildComparisonData() {
        const comparisonData = [];
        const periodInfo = {}; // Stocker les infos des périodes pour les en-têtes
        
        // Pour chaque PM sélectionné
        this.selectedPMs.forEach(pm => {
            const periods = Object.keys(this.pmData[pm]).sort();
            if (periods.length < 2) return; // Besoin d'au moins 2 périodes
            
            const oldestPeriod = periods[0];
            const newestPeriod = periods[periods.length - 1];
            
            // Stocker les dates pour les en-têtes
            if (!periodInfo.oldestDate && this.pmData[pm][oldestPeriod].length > 0) {
                periodInfo.oldestDate = this.pmData[pm][oldestPeriod][0].date;
            }
            if (!periodInfo.newestDate && this.pmData[pm][newestPeriod].length > 0) {
                periodInfo.newestDate = this.pmData[pm][newestPeriod][0].date;
            }
            
            // Grouper par famille et code d'erreur
            const errorCounts = {};
            
            // Compter pour la période la plus ancienne
            this.pmData[pm][oldestPeriod].forEach(record => {
                if (!this.selectedFamilies.includes(record.errorFamily)) return;
                
                const key = `${record.errorFamily}_${record.errorCode}`;
                if (!errorCounts[key]) {
                    errorCounts[key] = {
                        pm: pm,
                        family: record.errorFamily,
                        errorCode: record.errorCode,
                        description: record.fullRow[4] || 'N/A', // Colonne E (NATURE_ERREUR)
                        dsp: record.dsp || 'N/A', // DSP depuis le fichier de synthèse
                        oldCount: 0,
                        newCount: 0
                    };
                }
                errorCounts[key].oldCount++;
            });
            
            // Compter pour la période la plus récente
            this.pmData[pm][newestPeriod].forEach(record => {
                if (!this.selectedFamilies.includes(record.errorFamily)) return;
                
                const key = `${record.errorFamily}_${record.errorCode}`;
                if (!errorCounts[key]) {
                    errorCounts[key] = {
                        pm: pm,
                        family: record.errorFamily,
                        errorCode: record.errorCode,
                        description: record.fullRow[4] || 'N/A',
                        dsp: record.dsp || 'N/A', // DSP depuis le fichier de synthèse
                        oldCount: 0,
                        newCount: 0
                    };
                }
                errorCounts[key].newCount++;
            });
            
            // Calculer l'évolution
            Object.values(errorCounts).forEach(error => {
                const evolution = error.newCount - error.oldCount;
                const evolutionPercent = error.oldCount > 0 ? 
                    ((evolution / error.oldCount) * 100).toFixed(1) : 
                    (error.newCount > 0 ? '100' : '0');
                
                comparisonData.push({
                    ...error,
                    evolution: evolution,
                    evolutionPercent: evolutionPercent
                });
            });
        });
        
        // Ajouter les infos de période aux données pour l'affichage des en-têtes
        comparisonData.periodInfo = periodInfo;
        return comparisonData;
    }

    getEvolutionIcon(evolution) {
        if (evolution > 0) {
            return '<i class="bi bi-arrow-up text-danger"></i> DÉGRADATION';
        } else if (evolution < 0) {
            return '<i class="bi bi-arrow-down text-success"></i> AMÉLIORATION';
        } else {
            return '<i class="bi bi-arrow-right text-warning"></i> STABLE';
        }
    }

    populateChartsTab() {
        // Remplir les filtres
        this.populateChartFilters();
        
        // Détruire l'ancien graphique avant d'initialiser le nouveau
        if (this.mainChart) {
            this.mainChart.destroy();
            this.mainChart = null;
        }
        
        // Initialiser les graphiques
        this.initializeCharts();
        
        // Ajouter les listeners pour les filtres
        document.getElementById('chartPmFilter').addEventListener('change', () => this.updateCharts());
        document.getElementById('chartFamilyFilter').addEventListener('change', () => this.updateCharts());
        document.getElementById('chartErrorCodeFilter').addEventListener('change', () => this.updateCharts());
        document.getElementById('chartDisplayType').addEventListener('change', () => this.updateCharts());
    }

    populateChartFilters() {
        // Remplir le filtre PM
        const pmFilter = document.getElementById('chartPmFilter');
        pmFilter.innerHTML = '<option value="">Tous les PM</option>';
        this.selectedPMs.forEach(pm => {
            const option = document.createElement('option');
            option.value = pm;
            option.textContent = pm;
            pmFilter.appendChild(option);
        });

        // Remplir le filtre famille
        const familyFilter = document.getElementById('chartFamilyFilter');
        familyFilter.innerHTML = '<option value="">Toutes les familles</option>';
        this.selectedFamilies.forEach(family => {
            const option = document.createElement('option');
            option.value = family;
            option.textContent = family;
            familyFilter.appendChild(option);
        });

        // Remplir le filtre code d'erreur
        this.populateErrorCodeFilter();
    }

    populateErrorCodeFilter() {
        const errorCodeFilter = document.getElementById('chartErrorCodeFilter');
        errorCodeFilter.innerHTML = '<option value="">Tous les codes</option>';
        
        // Obtenir tous les codes d'erreur uniques des données filtrées
        const selectedPM = document.getElementById('chartPmFilter').value;
        const selectedFamily = document.getElementById('chartFamilyFilter').value;
        
        const filteredData = this.allData.filter(record => {
            let pmMatch = !selectedPM || record.refPM === selectedPM;
            let familyMatch = !selectedFamily || record.errorFamily === selectedFamily;
            return pmMatch && familyMatch && this.selectedPMs.includes(record.refPM);
        });
        
        const uniqueErrorCodes = [...new Set(filteredData.map(record => record.errorCode))].sort();
        
        uniqueErrorCodes.forEach(code => {
            const option = document.createElement('option');
            option.value = code;
            option.textContent = code;
            errorCodeFilter.appendChild(option);
        });
    }

    initializeCharts() {
        // Détruire le graphique existant s'il existe
        if (this.mainChart) {
            this.mainChart.destroy();
        }
        
        // Créer le graphique principal unique
        const ctx = document.getElementById('mainChart').getContext('2d');
        this.mainChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: [],
                datasets: []
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: false // Titre géré par HTML
                    },
                    legend: {
                        display: true,
                        position: 'top',
                        labels: {
                            usePointStyle: true,
                            pointStyle: 'rect',
                            font: {
                                size: 11
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Nombre d\'erreurs',
                            font: {
                                size: 11
                            }
                        },
                        ticks: {
                            font: {
                                size: 10
                            }
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Familles d\'erreurs',
                            font: {
                                size: 11
                            }
                        },
                        ticks: {
                            font: {
                                size: 10
                            },
                            maxRotation: 45
                        }
                    }
                },
                layout: {
                    padding: {
                        top: 10,
                        bottom: 10,
                        left: 10,
                        right: 10
                    }
                }
            }
        });

        // Mettre à jour avec les données actuelles
        this.updateCharts();
    }

    updateCharts() {
        const selectedPM = document.getElementById('chartPmFilter').value;
        const selectedFamily = document.getElementById('chartFamilyFilter').value;
        const selectedErrorCode = document.getElementById('chartErrorCodeFilter').value;
        const displayType = document.getElementById('chartDisplayType').value;

        // Mettre à jour le filtre des codes d'erreur si PM ou famille ont changé
        this.populateErrorCodeFilter();

        // Filtrer les données
        const filteredData = this.allData.filter(record => {
            let pmMatch = !selectedPM || record.refPM === selectedPM;
            let familyMatch = !selectedFamily || record.errorFamily === selectedFamily;
            let codeMatch = !selectedErrorCode || record.errorCode === selectedErrorCode;
            return pmMatch && familyMatch && codeMatch && this.selectedPMs.includes(record.refPM);
        });

        if (displayType === 'comparison') {
            this.updateComparisonChart(filteredData);
        } else {
            this.updateEvolutionChart(filteredData);
        }
    }

    updateComparisonChart(filteredData) {
        // Obtenir toutes les périodes uniques triées
        const allPeriods = [...new Set(filteredData.map(record => this.getPeriodKey(record.date)))].sort();
        
        if (allPeriods.length === 0) {
            this.mainChart.data.labels = [];
            this.mainChart.data.datasets = [];
            this.mainChart.update();
            return;
        }
        
        // Organiser les données par famille et par période
        const familyPeriodData = {};
        
        filteredData.forEach(record => {
            const family = record.errorFamily;
            const period = this.getPeriodKey(record.date);
            
            if (!familyPeriodData[family]) {
                familyPeriodData[family] = {};
            }
            if (!familyPeriodData[family][period]) {
                familyPeriodData[family][period] = 0;
            }
            familyPeriodData[family][period]++;
        });
        
        // Filtrer les familles qui ont au moins 1 erreur dans une période
        const validFamilies = Object.keys(familyPeriodData).filter(family => {
            return allPeriods.some(period => familyPeriodData[family][period] > 0);
        });
        
        // Ordonner par nombre total d'erreurs décroissant
        validFamilies.sort((a, b) => {
            const totalA = allPeriods.reduce((sum, period) => sum + (familyPeriodData[a][period] || 0), 0);
            const totalB = allPeriods.reduce((sum, period) => sum + (familyPeriodData[b][period] || 0), 0);
            return totalB - totalA;
        });
        
        // Définir les couleurs pour chaque période
        const periodColors = this.getPeriodColors(allPeriods);
        
        // Créer les datasets pour chaque période
        const datasets = allPeriods.map((period, index) => {
            const data = validFamilies.map(family => familyPeriodData[family][period] || 0);
            const color = periodColors[index];
            const date = this.formatDateFromPeriod(period);
            
            return {
                label: `${date}`,
                data: data,
                backgroundColor: color.background,
                borderColor: color.border,
                borderWidth: 1,
                barThickness: 20,
                categoryPercentage: 0.8,
                barPercentage: 0.9
            };
        });
        
        // Mettre à jour le titre
        const selectedPM = document.getElementById('chartPmFilter').value;
        const selectedFamily = document.getElementById('chartFamilyFilter').value;
        const selectedErrorCode = document.getElementById('chartErrorCodeFilter').value;
        
        let titleText = 'Comparaison des erreurs par famille';
        if (selectedPM) titleText += ` - PM: ${selectedPM}`;
        if (selectedErrorCode) titleText += ` - Code: ${selectedErrorCode}`;
        else if (selectedFamily) titleText += ` - Famille: ${selectedFamily}`;
        
        document.getElementById('chartTitle').textContent = titleText;
        
        // Mettre à jour le graphique
        this.mainChart.config.type = 'bar';
        this.mainChart.data.labels = validFamilies;
        this.mainChart.data.datasets = datasets;
        this.mainChart.update();
    }
    
    updateEvolutionChart(filteredData) {
        // Graphique d'évolution temporelle
        const periodData = {};
        
        filteredData.forEach(record => {
            const period = this.getPeriodKey(record.date);
            if (!periodData[period]) {
                periodData[period] = 0;
            }
            periodData[period]++;
        });

        const periods = Object.keys(periodData).sort();
        const counts = periods.map(period => periodData[period]);
        
        // Mettre à jour le titre
        const selectedPM = document.getElementById('chartPmFilter').value;
        const selectedFamily = document.getElementById('chartFamilyFilter').value;
        const selectedErrorCode = document.getElementById('chartErrorCodeFilter').value;
        let titleText = 'Évolution temporelle des erreurs';
        if (selectedPM) titleText += ` - PM: ${selectedPM}`;
        if (selectedErrorCode) titleText += ` - Code: ${selectedErrorCode}`;
        else if (selectedFamily) titleText += ` - Famille: ${selectedFamily}`;
        document.getElementById('chartTitle').textContent = titleText;

        // Convertir en graphique en ligne pour l'évolution
        this.mainChart.config.type = 'line';
        this.mainChart.data.labels = periods;
        this.mainChart.data.datasets = [{
            label: 'Nombre d\'erreurs',
            data: counts,
            borderColor: 'rgba(75, 192, 192, 1)',
            backgroundColor: 'rgba(75, 192, 192, 0.2)',
            tension: 0.1,
            fill: true
        }];
        this.mainChart.update();
    }
    
    countErrorsByCategory(records, categoryField, categoryName) {
        if (categoryName === 'Total') {
            return records.length;
        }
        return records.filter(record => record[categoryField] === categoryName).length;
    }
    
    getPeriodColors(periods) {
        // Couleurs spécifiques : Gris pour première, Bleu pour dernière, autres couleurs pour intermédiaires
        const colors = [
            { background: 'rgba(128, 128, 128, 0.7)', border: 'rgba(128, 128, 128, 1)' }, // Gris pour ancienne
            { background: 'rgba(75, 192, 192, 0.7)', border: 'rgba(75, 192, 192, 1)' },   // Vert pour intermédiaire
            { background: 'rgba(255, 206, 86, 0.7)', border: 'rgba(255, 206, 86, 1)' },   // Jaune pour intermédiaire  
            { background: 'rgba(153, 102, 255, 0.7)', border: 'rgba(153, 102, 255, 1)' }, // Violet pour intermédiaire
            { background: 'rgba(54, 162, 235, 0.7)', border: 'rgba(54, 162, 235, 1)' }    // Bleu pour nouvelle
        ];
        
        if (periods.length === 1) {
            return [colors[4]]; // Bleu si une seule période
        } else if (periods.length === 2) {
            return [colors[0], colors[4]]; // Gris + Bleu
        } else {
            // Pour 3+ périodes : Gris, couleurs intermédiaires, Bleu
            const result = [colors[0]]; // Première = Gris
            for (let i = 1; i < periods.length - 1; i++) {
                result.push(colors[Math.min(i, colors.length - 2)]);
            }
            result.push(colors[4]); // Dernière = Bleu
            return result;
        }
    }
    
    formatDateFromPeriod(period) {
        // Convertir "2025-04-23" en "23/04/2025"
        const [year, month, day] = period.split('-');
        return `${day}/${month}/${year}`;
    }

    getChartColor(index, alpha = 1) {
        const colors = [
            `rgba(255, 99, 132, ${alpha})`,
            `rgba(54, 162, 235, ${alpha})`,
            `rgba(255, 205, 86, ${alpha})`,
            `rgba(75, 192, 192, ${alpha})`,
            `rgba(153, 102, 255, ${alpha})`,
            `rgba(255, 159, 64, ${alpha})`,
            `rgba(199, 199, 199, ${alpha})`,
            `rgba(83, 102, 255, ${alpha})`
        ];
        return colors[index % colors.length];
    }

    showProgress(show, text = '') {
        const section = document.getElementById('progressSection');
        if (show) {
            section.style.display = 'block';
            document.getElementById('progressText').textContent = text;
        } else {
            section.style.display = 'none';
        }
    }

    updateProgress(percent, details = '') {
        document.getElementById('progressBar').style.width = percent + '%';
        document.getElementById('progressDetails').textContent = details;
    }

    updateStatus(message) {
        console.log('Status:', message);
    }

    showError(message) {
        document.getElementById('errorMessage').textContent = message;
        const modal = new bootstrap.Modal(document.getElementById('errorModal'));
        modal.show();
    }

    reset() {
        // Détruire le graphique existant
        if (this.mainChart) {
            this.mainChart.destroy();
            this.mainChart = null;
        }
        
        // Réinitialiser toutes les données
        this.zipFiles = [];
        this.allData = [];
        this.pmData = {};
        this.selectedPMs = [];
        this.selectedFamilies = [];
        this.csvHeaders = null; // Réinitialiser aussi les en-têtes
        
        // Réinitialiser l'interface
        document.getElementById('zipFiles').value = '';
        document.getElementById('referenceFile').value = '';
        document.getElementById('analyzeBtn').disabled = true;
        document.getElementById('analyzeAndVisualize').disabled = true;
        document.getElementById('generateReportBtn').disabled = true;
        document.getElementById('filterSection').style.display = 'none';
        document.getElementById('progressSection').style.display = 'none';
        document.getElementById('resultsSection').style.display = 'none';
        
        this.updateStatus('Application réinitialisée');
    }
}

// Initialiser l'application au chargement de la page
document.addEventListener('DOMContentLoaded', () => {
    window.graceApp = new GraceAnalyzer();
});