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
            

            var ac_power_maps = globalContext.get("map").ac_power;
            var multimeter_maps = globalContext.get("map").multimeter;
            var communication_maps = globalContext.get("map").communication;
            var relay_maps = globalContext.get("map").relay;
            var gpio_maps = globalContext.get("map").gpio;
            var mux_maps = globalContext.get("map").mux;

            const workbook = new Excel.Workbook();
            const worksheet = workbook.addWorksheet("JIG Mapeamento");

            worksheet.getCell('B2').value = "- AC POWER MAPPING -";
            worksheet.getCell('B2').fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFA8072'} };
            worksheet.mergeCells('B2:D2');
            
            worksheet.getColumn("B").width = 20;
            worksheet.getColumn("B").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("C").width = 10;
            worksheet.getColumn("C").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("D").width = 20;
            worksheet.getColumn("D").alignment = { vertical: 'middle', horizontal: 'center' };
            
            var rows_to_skip = 3; //Primeira table (ideia intuitiva é saltar o numero de linhas da tebela para gerar um novo header para cada slot de mapeamento)
            for(var currentMap of ac_power_maps){
                if(currentMap.length > 0){
                    worksheet.addTable({
                        name: `AC_POWER_${rows_to_skip}`,
                        ref: `B${rows_to_skip}`,
                        headerRow: true,
                        totalsRow: false,
                        columns: [
                            { name: 'Feature', key: 'feat', width: 20  },
                            { name: 'Pin', key: 'pin', width: 20 },
                            { name: '(TP or Connector)', key: 'board', width: 20 },
                        ],
                        rows: []
                    });
                    var AC_POWER = worksheet.getTable(`AC_POWER_${rows_to_skip}`);
                    for(var row of currentMap){
                        if(row.feat != ""){
                            AC_POWER.addRow([row.feat, row.pin, row.board]);
                            AC_POWER.commit();
                        }
                    }
                    rows_to_skip += 12;
                }
            }

            worksheet.getCell('G2').value = "- MULTIMETER MAPPING -";
            worksheet.getCell('G2').fill = { type: 'pattern', pattern: 'solid', fgColor:{argb:'FF32CD32'} };
            worksheet.mergeCells('G2:I2');

            worksheet.getColumn("G").width = 20;
            worksheet.getColumn("G").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("H").width = 10;
            worksheet.getColumn("H").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("I").width = 20;
            worksheet.getColumn("I").alignment = { vertical: 'middle', horizontal: 'center' };

            rows_to_skip = 3;
            for(var currentMap of multimeter_maps){
                if(currentMap.length > 0){
                    worksheet.addTable({
                        name: `MULTIMETER_${rows_to_skip}`,
                        ref: `G${rows_to_skip}`,
                        headerRow: true,
                        totalsRow: false,
                        columns: [
                            { name: 'Feature', key: 'feat', width: 20  },
                            { name: 'Pin', key: 'pin', width: 20 },
                            { name: '(TP or Connector)', key: 'board', width: 20 },
                        ],
                        rows: []
                    });
                    var MULTIMETER = worksheet.getTable(`MULTIMETER_${rows_to_skip}`);
                    currentMap.forEach( (item) => {
                        if(item.pin !== ""){
                            MULTIMETER.addRow([item.feat, item.pin, item.board]);
                        }else {
                            MULTIMETER.addRow([,item.feat,]);
                        }
                        MULTIMETER.commit();
                    });
                    rows_to_skip += 35;
                }
            }

            worksheet.getCell('L2').value = "- COMMUNICATION MAPPING -";
            worksheet.getCell('L2').fill = { type: 'pattern', pattern: 'solid', fgColor:{argb:'FF0080FF'} };
            worksheet.mergeCells('L2:N2');

            worksheet.getColumn("L").width = 20;
            worksheet.getColumn("L").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("M").width = 10;
            worksheet.getColumn("M").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("N").width = 20;
            worksheet.getColumn("N").alignment = { vertical: 'middle', horizontal: 'center' };

            rows_to_skip = 3;
            for(var currentMap of communication_maps){
                if(currentMap.length > 0){
                    worksheet.addTable({
                        name: `COMMUNICATION_${rows_to_skip}`,
                        ref: `L${rows_to_skip}`,
                        headerRow: true,
                        totalsRow: false,
                        columns: [
                            { name: 'Feature', key: 'feat', width: 20  },
                            { name: 'Pin', key: 'pin', width: 20 },
                            { name: '(TP or Connector)', key: 'board', width: 20 },
                        ],
                        rows: []
                    });
                    var COMMUNICATION = worksheet.getTable(`COMMUNICATION_${rows_to_skip}`);
                    currentMap.forEach( (item) => {
                        if(item.pin !== ""){
                            COMMUNICATION.addRow([item.feat, item.pin, item.board]);
                        }else {
                            COMMUNICATION.addRow([,item.feat,]);
                        }
                        COMMUNICATION.commit();
                    });
                    rows_to_skip += 41;
                }
                
            }

            worksheet.getCell('Q2').value = "- RELAY MAPPING -";
            worksheet.getCell('Q2').fill = { type: 'pattern', pattern: 'solid', fgColor:{argb:'FF808080'} };
            worksheet.mergeCells('Q2:S2');
            worksheet.getColumn("Q").width = 20;
            worksheet.getColumn("Q").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("R").width = 10;
            worksheet.getColumn("R").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("S").width = 20;
            worksheet.getColumn("S").alignment = { vertical: 'middle', horizontal: 'center' };

            rows_to_skip = 3;
            for(var currentMap of relay_maps){
                if(currentMap.length > 0){
                    worksheet.addTable({
                        name: `RELAY_${rows_to_skip}`,
                        ref: `Q${rows_to_skip}`,
                        headerRow: true,
                        totalsRow: false,
                        columns: [
                            { name: 'Feature', key: 'feat', width: 20  },
                            { name: 'Pin', key: 'pin', width: 20 },
                            { name: '(TP or Connector)', key: 'board', width: 20 },
                        ],
                        rows: []
                    });
                    var RELAY = worksheet.getTable(`RELAY_${rows_to_skip}`);
                    currentMap.forEach( (item) => {
                        if(item.pin !== ""){
                            RELAY.addRow([item.feat, item.pin, item.board]);
                        }else {
                            RELAY.addRow([,item.feat,]);
                        }
                        RELAY.commit();
                    });
                    rows_to_skip += 18;
                }
            }

            worksheet.getCell('V2').value = "- GPIO MAPPING -";
            worksheet.getCell('V2').fill = { type: 'pattern', pattern: 'solid', fgColor:{argb:'FFFFFF00'} };
            worksheet.mergeCells('V2:X2');

            worksheet.getColumn("V").width = 20;
            worksheet.getColumn("V").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("W").width = 10;
            worksheet.getColumn("W").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("X").width = 20;
            worksheet.getColumn("X").alignment = { vertical: 'middle', horizontal: 'center' };

            rows_to_skip = 3;
            for(var currentMap of gpio_maps){
                if(currentMap.length > 0){
                    worksheet.addTable({
                        name: `GPIO_${rows_to_skip}`,
                        ref: `V${rows_to_skip}`,
                        headerRow: true,
                        totalsRow: false,
                        columns: [
                            { name: 'Feature', key: 'feat', width: 20  },
                            { name: 'Pin', key: 'pin', width: 20 },
                            { name: '(TP or Connector)', key: 'board', width: 20 },
                        ],
                        rows: []
                    });

                    var GPIO = worksheet.getTable(`GPIO_${rows_to_skip}`);
                    currentMap.forEach( (item) => {
                        if(item.pin !== ""){
                            GPIO.addRow([item.feat, item.pin, item.board]);
                        }else {
                            GPIO.addRow([,item.feat,]);
                        }
                        GPIO.commit();
                    });
                    rows_to_skip += 27;
                }
                    
            }

            worksheet.getCell('AA2').value = "- MUX MAPPING -"
            worksheet.mergeCells('AA2:AC2');
            worksheet.getColumn("AA").width = 20;
            worksheet.getColumn("AA").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("AB").width = 20;
            worksheet.getColumn("AB").alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn("AC").width = 20;
            worksheet.getColumn("AC").alignment = { vertical: 'middle', horizontal: 'center' };

            rows_to_skip = 3;
            for(var currentMap of mux_maps){
                if(currentMap.length > 0){
                    worksheet.addTable({
                        name: `MUX_${rows_to_skip}`,
                        ref: `AA${rows_to_skip}`,
                        headerRow: true,
                        totalsRow: false,
                        columns: [
                            { name: 'Feature', key: 'feat', width: 20  },
                            { name: 'Pin', key: 'pin', width: 20 },
                            { name: '(TP or Connector)', key: 'board', width: 20 },
                        ],
                        rows: []
                    });
                    var MUX = worksheet.getTable(`MUX_${rows_to_skip}`);
                    currentMap.forEach( (item) => {
                        if(item.pin !== ""){
                            MUX.addRow([item.feat, item.pin, item.board]);
                        }else {
                            MUX.addRow([,item.feat,]);
                        }
                        MUX.commit();
                    });
                    rows_to_skip += 42;
                }    
            };

            await workbook.xlsx.writeFile(directory+'/jig_map.xlsx')
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