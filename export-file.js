module.exports = function(RED) {

    "use strict";
    var fs = require("fs-extra");
    var path = require("path");

    function encode(data) {
        return Buffer.from(data);
    }

    function ExportFileNode(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        node.status({fill:"green", shape:"ring", text:"available"});
        node.filename = config.filename;
        node.directory = config.directory;
        node.createDir = config.createDir || false;

        node.wstream = null;
        node.msgQueue = [];
        node.closing = false;
        node.closeCallback = null;

        function processMsg(msg,nodeSend, done) {
            var filename = node.directory + "/" + node.filename + ".json"
            if (node.directory === "") {
                node.status({fill:"red",shape:"dot",text:"Error"});
                node.warn(RED._("Directory empty"));
                done();
            }
            if (node.filename === "") {
                node.status({fill:"red",shape:"dot",text:"Error"});
                node.warn(RED._("Filename empty"));
                done();
            }
            else{ 
                if (msg.hasOwnProperty("payload") && (typeof msg.payload !== "undefined")) {
                    var dir = path.dirname(filename);
                    var data = msg.payload;
                    if (node.createDir) {
                        try {
                            fs.ensureDirSync(dir);
                        } catch(err) {
                            node.status({fill:"red",shape:"dot",text:"Error"});
                            node.error(RED._("Fail to create directory",{error:err.toString()}),msg);
                            done();
                            return;
                        }
                    }
                    data = JSON.stringify(data);

                    var buf = encode(data);

                    var wstream = fs.createWriteStream(filename, { encoding:'binary', flags:'w', autoClose:true });
                    node.wstream = wstream;
                    wstream.on("error", function(err) {
                        node.status({fill:"red",shape:"dot",text:"Error"});
                        node.error(RED._("Error to save file",{error:err.toString()}),msg);
                        done();
                    });
                    wstream.on("open", function() {
                        wstream.end(buf, function() {
                            nodeSend(msg);
                            done();
                        });
                    })
                    return;
                }
                else {
                    done();
                }
            }
        }

        function processQueue(queue) {
            var event = queue[0];

            processMsg(event.msg, event.send, function() {
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

        node.on('input', function(msg, send, done) {
            var globalContext = this.context().global;
            var exportMode = globalContext.get("exportMode");
            var exportFile = globalContext.get("exportFile");
            var exportFileEmpty = {
                "slots": [
                    {
                        "jig_test": [],
                        "jig_error": []
                    },
                    {
                        "jig_test": [],
                        "jig_error": []
                    },
                    {
                        "jig_test": [],
                        "jig_error": []
                    },
                    {
                        "jig_test": [],
                        "jig_error": []
                    },
                ]
            }
            globalContext.set("exportFile", exportFileEmpty)

            if(exportMode){
                node.status({fill:"green",shape:"dot",text:"Generating"});
                var file = {
                    payload: exportFile
                }
                var msgQueue = node.msgQueue;

                msgQueue.push(
                    {
                        msg: file,
                        send: send,
                        done: done
                    }
                )
                if (msgQueue.length > 1) {
                    return;
                }

                try {
                    processQueue(msgQueue);
                }
                catch(e) {
                    node.status({fill:"red",shape:"dot",text:"Error"});
                    node.msgQueue = [];
                    if (node.closing) {
                        closeNode();
                    }
                    throw e;                    
                }
                node.status({fill:"green",shape:"dot",text:"Json generated"});
                done();
            }
            else{
                node.status({fill:"yellow",shape:"dot",text:"Warn: Export mode is FALSE"});
                done();
            }
        });
    }

    RED.nodes.registerType("export-file", ExportFileNode);
}