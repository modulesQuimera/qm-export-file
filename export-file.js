module.exports = function (RED) {

    "use strict";
    var fs = require("fs-extra");
    var path = require("path");
    var Excel = require("exceljs");

    function encode(data) {
        return Buffer.from(data);
    }

    function export_fileNode(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        node.status({ fill: "green", shape: "ring", text: "available" });
        node.filename = config.filename;
        node.directory = config.directory;
        node.createDir = config.createDir || false;

        node.wstream = null;
        node.msgQueue = [];
        node.closing = false;
        node.closeCallback = null;

        function processMsg(msg, nodeSend, done) {
            var filename = node.directory + "/" + node.filename + ".json";
            if (node.directory === "") {
                node.status({ fill: "red", shape: "dot", text: "Error" });
                node.warn(RED._("Directory empty"));
                done();
            }
            if (node.filename === "") {
                node.status({ fill: "red", shape: "dot", text: "Error" });
                node.warn(RED._("Filename empty"));
                done();
            }
            else {
                if (msg.hasOwnProperty("payload") && (typeof msg.payload !== "undefined")) {
                    var dir = path.dirname(filename);
                    var data = msg.payload;
                    if (node.createDir) {
                        try {
                            fs.ensureDirSync(dir);
                        } catch (err) {
                            node.status({ fill: "red", shape: "dot", text: "Error" });
                            node.error(RED._("Fail to create directory", { error: err.toString() }), msg);
                            done();
                            return;
                        }
                    }
                    data = JSON.stringify(data);

                    var buf = encode(data);

                    var wstream = fs.createWriteStream(filename, { encoding: 'binary', flags: 'w', autoClose: true });
                    node.wstream = wstream;
                    wstream.on("error", function (err) {
                        node.status({ fill: "red", shape: "dot", text: "Error" });
                        node.error(RED._("Error to save file", { error: err.toString() }), msg);
                        done();
                    });
                    wstream.on("open", function () {
                        wstream.end(buf, function () {
                            nodeSend(msg);
                            done();
                        });
                    });
                    return;
                }
                else {
                    done();
                }
            }
        }

        function processQueue(queue) {
            var event = queue[0];

            processMsg(event.msg, event.send, function () {
                event.done();
                queue.shift();
                if (queue.length > 0) {
                    processQueue(queue);
                }
                else if (node.closing) {
                    closeNode();
                }
            });
        }

        async function generateFILE(globalContext, directory, done){
            
            var file = globalContext.get("exportFile");
            file = JSON.stringify(file);
            
            const workbook = new Excel.Workbook();
            const worksheet = workbook.addWorksheet("JIG Mapeamento");

            var MODULES_MAPS = [
                { type: 'AC_power_source_virtual_V1_0', feat: 'B', pin: 'C', board: 'D', skip: 12, maps: globalContext.get("map").ac_power, colorHeader: 'FFFA8072', titleHeader: "- AC POWER MAPPING -" },
                { type: 'multimeter_modular_V1_0', feat: 'G', pin: 'H', board: 'I', skip: 35, maps: globalContext.get("map").multimeter, colorHeader: 'FF32CD32', titleHeader: "- MULTIMETER MAPPING -" },
                { type: 'communication_modular_V1_0', feat: 'L', pin: 'M', board: 'N', skip: 41, maps: globalContext.get("map").communication, colorHeader: 'FF0080FF', titleHeader: "- COMMUNICATION MAPPING -" },
                { type: 'relay_modular_V1_0', feat: 'Q', pin: 'R', board: 'S', skip: 18, maps: globalContext.get("map").relay, colorHeader: 'FF808080', titleHeader: "- RELAY MAPPING -" },
                { type: 'GPIO_modular_V1_0', feat: 'V', pin: 'W', board: 'X', skip: 27, maps: globalContext.get("map").gpio, colorHeader: 'FFFFFF00', titleHeader: "- GPIO MAPPING -" },
                { type: 'mux_modular_V1_0', feat: 'AA', pin: 'AB', board: 'AC', skip: 42, maps: globalContext.get("map").mux, colorHeader: 'FFFFFFFF', titleHeader: "- MUX MAPPING -" },
            ]

            for(var CURRENT_MODULE of MODULES_MAPS){

                worksheet.getCell(`${CURRENT_MODULE.feat}2`).value = CURRENT_MODULE.titleHeader;
                worksheet.getCell(`${CURRENT_MODULE.feat}2`).fill = { type: 'pattern', pattern:'solid', fgColor:{ argb: CURRENT_MODULE.colorHeader } };
                worksheet.mergeCells(`${CURRENT_MODULE.feat}2:${CURRENT_MODULE.board}2`);
                
                worksheet.getColumn(CURRENT_MODULE.feat).width = 20;
                worksheet.getColumn(CURRENT_MODULE.feat).alignment = { vertical: 'middle', horizontal: 'center' };
                worksheet.getColumn(CURRENT_MODULE.pin).width = 20;
                worksheet.getColumn(CURRENT_MODULE.pin).alignment = { vertical: 'middle', horizontal: 'center' };
                worksheet.getColumn(CURRENT_MODULE.board).width = 20;
                worksheet.getColumn(CURRENT_MODULE.board).alignment = { vertical: 'middle', horizontal: 'center' };
                
                var rows_to_skip = 3; //Primeira table (ideia intuitiva é saltar o numero de linhas da tebela para gerar um novo header para cada slot de mapeamento)
                if(file.indexOf(CURRENT_MODULE.type) !== -1){ // verifica se o mapeamento está sendo utlizado no flow atual.
                    for(var currentMap of CURRENT_MODULE.maps){
                        if(currentMap.length > 0){
                            var CURRENT_TABLE = worksheet.addTable({
                                name: `${CURRENT_MODULE.type}_${rows_to_skip}`,
                                ref: `${CURRENT_MODULE.feat}${rows_to_skip}`,
                                headerRow: true,
                                totalsRow: false,
                                columns: [
                                    { name: 'Feature', key: 'feat', width: 20  },
                                    { name: 'Pin', key: 'pin', width: 20 },
                                    { name: '(TP or Connector)', key: 'board', width: 20 },
                                ],
                                rows: []
                            });
                            
                            currentMap.forEach( (item) => {
                                if(item.pin !== ""){
                                    CURRENT_TABLE.addRow([item.feat, item.pin, item.board]);
                                }else {
                                    CURRENT_TABLE.addRow([,item.feat,]);
                                }
                                CURRENT_TABLE.commit();
                            });
                            rows_to_skip += CURRENT_MODULE.skip;
                        }
                    }
                }
            }

            await workbook.xlsx.writeFile(`${directory}/${node.filename}_jig_map.xlsx`)
            .then(() => {
                console.log('The JIG MAPPING file was written successfully.')
                node.status({ fill: "green", shape: "dot", text: "The JIG MAPPING file was written successfully." });
            })
            .catch(() => {
                node.status({ fill: "red", shape: "dot", text: "The JIG MAPPING file are in use on another program." });
                console.log('The JIG MAPPING file are in use on another program.');
            });

        }

        node.on('input', function (msg, send, done) {

            var globalContext = this.context().global;
            var exportFile = globalContext.get("exportFile");

            var quantidade = globalContext.get("export_file") + 1;
            globalContext.set("export_file", quantidade);


            generateFILE(globalContext, this.directory, done);

            node.status({ fill: "green", shape: "dot", text: "Generating" });
            var file = {
                payload: exportFile
            };
            var msgQueue = node.msgQueue;

            msgQueue.push(
                {
                    msg: file,
                    send: send,
                    done: done
                }
            );
            if (msgQueue.length > 1) {
                return;
            }

            try {
                processQueue(msgQueue);
            }
            catch (e) {
                node.status({ fill: "red", shape: "dot", text: "Error" });
                node.msgQueue = [];
                if (node.closing) {
                    closeNode();
                }
                throw e;
            }
            node.status({ fill: "green", shape: "dot", text: "Json generated" });
            done();
        });
    }

    RED.nodes.registerType("export-file", export_fileNode);
};