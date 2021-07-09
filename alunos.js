"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
exports.__esModule = true;
var google_spreadsheet_1 = require("google-spreadsheet");
var Aluno = /** @class */ (function () {
    function Aluno(matricula, nome, faltas, P1, P2, P3, situacao, Naf) {
        this.matricula = matricula;
        this.nome = nome;
        this.faltas = faltas;
        this.P1 = P1;
        this.P2 = P2;
        this.P3 = P3;
        this.situacao = situacao;
        this.Naf = Naf;
        //Calculando a média...
        this.media = Number(((P1 + P2 + P3) / 3).toFixed(2));
    }
    return Aluno;
}());
var alunosLidos;
var doc = new google_spreadsheet_1.GoogleSpreadsheet('1j0p6AO0gkDe1mL8y_-fiocti60x6oJ1G4urTL_VmzyM');
(function () {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, principal()];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}());
function principal() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, LerAlunos()];
                case 1:
                    _a.sent();
                    alunosLidos.forEach(function (aluno) {
                        calculaSituacaoAluno(aluno);
                        calculaNafAluno(aluno);
                    });
                    salvaAlunos();
                    return [2 /*return*/];
            }
        });
    });
}
function calculaSituacaoAluno(aluno) {
    // Se faltas forem maior do que 25% das 60 aulas, a situação do aluno será 'Reprovado por falta'
    if (aluno.faltas > 15) {
        aluno.situacao = 'Reprovado por Falta';
        return;
    }
    // A média das provas P1, P2 e P3 menor do que 5, situação é 'Reprovado por nota'
    if (aluno.media < 50) {
        aluno.situacao = 'Reprovado por nota';
        return;
    }
    // A média das provas P1, P2 e P3 maior ou igual a 5 e menor do que 7, a situação é 'Exame final'
    if (aluno.media >= 50 && aluno.media < 70) {
        aluno.situacao = 'Exame Final';
        return;
    }
    // A média das provas P1, P2 e P3 maior ou igual 7, a situação é 'Aprovado'
    aluno.situacao = 'Aprovado';
}
function calculaNafAluno(aluno) {
    // A Naf é uma nota mínima, ou meta do aluno, na prova final.
    // A relação deve ser respeitada: 5 <= (m + naf)/2
    //   - Isso significa, que temos que encontrar o valor mínimo de naf. Para isso, faremos:
    //   - naf >= 10 - m, ou seja, o valor mínimo para a naf é (10 - m)
    // Caso a situação do aluno seja diferente de 'Exame final', preencher o Naf com 0
    if (aluno.situacao != 'Exame Final') {
        aluno.Naf = 0;
        return;
    }
    // Arredondar o valor para o próximo número inteiro. Ou seja, o Naf não pode ser decimal.
    aluno.Naf = Math.round(100 - aluno.media);
}
function LerAlunos() {
    return __awaiter(this, void 0, void 0, function () {
        var vetorAluno, sheet, linhas, contaAlunos, i, matricula, nome, faltas, p1, p2, p3, situacao, Naf;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, accessSpreadsheet()];
                case 1:
                    _a.sent();
                    vetorAluno = [];
                    sheet = doc.sheetsByIndex[0];
                    console.log(sheet.title);
                    return [4 /*yield*/, sheet.getRows()];
                case 2:
                    linhas = _a.sent();
                    contaAlunos = linhas.length;
                    return [4 /*yield*/, sheet.loadCells('A1:H500')];
                case 3:
                    _a.sent(); // loads a range of cells
                    for (i = 3; i <= contaAlunos; i++) {
                        matricula = Number(sheet.getCell(i, 0).value);
                        nome = sheet.getCell(i, 1).value;
                        faltas = Number(sheet.getCell(i, 2).value);
                        p1 = Number(sheet.getCell(i, 3).value);
                        p2 = Number(sheet.getCell(i, 4).value);
                        p3 = Number(sheet.getCell(i, 5).value);
                        situacao = sheet.getCell(i, 6).value;
                        Naf = Number(sheet.getCell(i, 7).value);
                        vetorAluno.push(new Aluno(matricula, nome, faltas, p1, p2, p3, situacao, Naf));
                    }
                    alunosLidos = vetorAluno;
                    return [2 /*return*/];
            }
        });
    });
}
function accessSpreadsheet() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, doc.useServiceAccountAuth({
                        client_email: "desafiotunts@silicon-dialect-319119.iam.gserviceaccount.com",
                        private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCxM+zbg/0m2m6r\n6cVhfM0fXzIVfb8Wgm7DN59YorvHrKr+WjkmH6UMKLbdiiabUPsZ1OWh3KFxE6tz\nxd3wymV3qrg7hpo4JPQv1/TZtFhI3YGUCCoMKdZK0FAO26yb+FEyVNi7gtmsLFy5\nxvPCt5ZPBVZ7NdEi2cDsCLORUNpRTEbkcZL9g9UsVG7F3B3ILNwZfzpFMCf3HUMF\nqDLconTk1S61WL2a0Ao/y4JNwqhApD+RuzuRckeIcMrItM3pGA1SZIbEgFYUrSU6\nyro9nZCh4VDahuINKqf7aqkEkhAECDhXO8d42pM79VpuXc50TVLFXbAmjMDbXfRc\nc1V6xSylAgMBAAECggEAAs/9quWXOit5BTPSelUg66YMT/NXzz4mZDp48M8IZvIS\ndrNXjoeDDESF09H9AeBmU9x2q+Z/Y3bkn5Oc4v+mCPZbMoqqM55BMpSIDNUuxN1L\nJgkpIwSeTkAB0p029J9w086CmsXyuX0CGGq+ygY1la8n1mXIxM0ATF6+UYifJm97\nixJGmWp3cJnd4FeA51cSvO/+NBd1AcFDNxLqbC+hD1+JOgNLaYcOACJ6iMA/5zuH\nJYoscTunN7fQIEVj7F5ws69WOuYDlGwCB0RtI6fRluPV4R8VlvHqrCLlr7GqIogR\nHSPeJlKueEUM/8uXBW/TqfOYraALo13CRlV4jymmcwKBgQDlP+rwrLP+K9GW08lu\nmOl7WAFRshXAM8aaAM8L9nAmoMWyl/n8jpzdA5BR7fsTRjxO2Kvmewrx/A9x4KOc\nseoWu216kKKngBdC7ws7hUYKKKgrMAITHVi236tmlz75K1BRj8fv0F9cizG8uaNN\nZULHb0niU24xU7/MEgmle4CJ0wKBgQDF4UdXNB+TTcww8llCPEWdYR9IqVZE2Dui\n0wTWjw8sgbBBSAZV3hpFblITujKlD+A86289R0Zd8qydWL4Urzwjbws//icnBzZS\nX9FHfP9lpd7WwYoWwQQJyfAdpsm3rNG/rZLyEYgahbrlE3XspSrXr65U+BBDIgn2\nOX5YrPkspwKBgQDMQLm7u7QW1sXDY2kKMBV+vvdV6Zx1hewCFIxktSpRUFc0ezHR\nMuNSC7XPOYDFOIPNIEFwddpXpePA64v6tY3CuTWeyTUSlg6jpUXVvzWbIYRMDlca\n3r/HF9un6UPDTzMdqERUR8xfMOmco6167Kil9mLW0szQCDVPxhZwKWxp3wKBgGX9\ntlqhGAFBoRQ9ZXo5PJxgedJmzXtQhHRpFV5NgEGtWp5bNEC/6ISO1ykp2H6xTx+3\nLa/E8+TqdsPnAJoCtBmDW6YRJjb8hagxkNmq+Kx4sQG54aXWuHEfL27pD6FnJvkH\nyuyP0rnw4aK+xBJEE2/2MgHDlgY0HjRV7+Rey1OTAoGADR+YGnQp0VIKWkR2afGA\n5viPyyIHtYbiGjSjfJbefqH/8KlIgSL7IU4XL5KDdeu9pMv8cb82PuBxhG5fQu/1\nSbZgIbQTXFwOsdMyNp2xDz8Iptv9O3Ksc8EWfiw/22ylLfipk8xrHK+KnY3bDJtK\nfn8jM8h0CXDqk5bCwka5jCQ=\n-----END PRIVATE KEY-----\n"
                    })];
                case 1:
                    _a.sent();
                    return [4 /*yield*/, doc.loadInfo()];
                case 2:
                    _a.sent(); // loads document properties and worksheets
                    console.log(doc.title);
                    return [2 /*return*/];
            }
        });
    });
}
function salvaAlunos() {
    return __awaiter(this, void 0, void 0, function () {
        var sheet, linhas, contaAlunos, i, indiceAluno, situacao, Naf;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    sheet = doc.sheetsByIndex[0];
                    return [4 /*yield*/, sheet.getRows()];
                case 1:
                    linhas = _a.sent();
                    contaAlunos = linhas.length;
                    return [4 /*yield*/, sheet.loadCells('A1:H500')];
                case 2:
                    _a.sent(); // loads a range of cells
                    i = 3;
                    _a.label = 3;
                case 3:
                    if (!(i <= contaAlunos)) return [3 /*break*/, 6];
                    indiceAluno = i - 3;
                    situacao = sheet.getCell(i, 6);
                    Naf = sheet.getCell(i, 7);
                    situacao.value = alunosLidos[indiceAluno].situacao;
                    Naf.value = alunosLidos[indiceAluno].Naf;
                    return [4 /*yield*/, sheet.saveUpdatedCells()];
                case 4:
                    _a.sent(); // save all updates in one call
                    _a.label = 5;
                case 5:
                    i++;
                    return [3 /*break*/, 3];
                case 6: return [2 /*return*/];
            }
        });
    });
}
