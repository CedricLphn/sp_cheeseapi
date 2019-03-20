//     ___ _____
//    /\ (_)    \
//   /  \      (_,
//  _)  _\   _    \
// /   (_)\_( )____\  <LPHN>
// \_     /    _  _/
//   ) /\/  _ (o)(
//   \ \_) (o)   /
//    \/________/         mic
//
// Sharepoint API Rest JS
// http://github.com/cedricLPHN/sp_cheese_api

class SP {

    constructor(url = null) {
        this.config = {
            url: url,
            list: null
        }

        this.user = {
            context: null,
            currentuser: null,
            digest: null,
            data: null
        }

        return true;
    }

    getContext() {
        return new Promise(async (resolve, reject) => {
            var _this = this;
            if (_this.user.context == null) {
                $.ajax({
                    url: _this.config.url + '/_api/contextinfo',
                    method: "POST",
                    headers: {
                        "Accept": "application/json; odata=verbose"
                    },
                    dataType: 'json',
                    success: function (msg) {
                        _this.user.context = msg;
                        resolve(msg);
                        return _this.user.context;
                    },
                    error: function (err) {
                        reject(err);
                    }
                })

            } else {
                console.log("context déjà rempli");
                resolve(_this.user.context);
            }

        })

    }

    getDigest() {
        return new Promise((resolve, reject) => {
            this.getContext().then(async (digest) => { 
                console.log("getdigest ok");
                resolve(digest.d.GetContextWebInformation.FormDigestValue) 
            }).catch(async () => { 
               reject(false);
            });
        })

    }

    setUrl(url) {
        this.config.url = url;
    }

    getUrl() {
        return this.config.url;
    }

    setList(name) {
        this.config.list = name;
        return true;
    }

    getList() {
        return this.config.list;
    }

    GetItemTypeForListName(name) {
        let n = name.replace("_", "_x005f_");
        n = n.charAt(0).toUpperCase() + n.split(" ").join("").slice(1);
        return "SP.Data." + n + "ListItem";
    }

    _convertFiltersLine(object = {}) {
        let filters = '';
        if (object.length === 0 || object === undefined) {
            return filters
        }

        let count = 0;
        for (let index in object) {
            (count == 0) ? filters += '': filters += '&';
            if (index == "select") {
                filters += `$select=`
                $.each(object[index], (i, value) => {
                    if (i < object[index].length - 1) {
                        filters += `${value},`;
                    } else {
                        filters += value;
                    }
                });
            } else if (index == "filters") {
                filters += `$filter=`
                let j = 0;
                $.each(object[index], (i, value) => {
                    console.log("index", i);
                    if (i != 'custom') {
                        filters += `${i} eq '${value}' `
                        console.log(j);
                        if (count > 0 && j < Object.keys(object[index]).length - 1) {
                            filters += 'and ';
                        }
                    } else {
                        filters += value;
                    }
                    j++;
                });
            } else if (index == "orderby") {
                (count === 0) ? filters += `$orderby=${object[index]}`: filters += `$orderby=${object[index]}`;
            } else if (index == "top") {
                (count === 0) ? filters += `$top=${object[index]}`: filters += `$top=${object[index]}`;
            } else if(index == "expand") {
                filters += `$expand=${object[index]}`;
            }
            count++;

        }

        return filters;

    }

    fetchAll(args) {
        var _this = this;
        var listname = _this.getList();
        if(args.listName) { listname = args.listName; }

        return new Promise(async (resolve, reject) => {
            console.log("async", args.async);
            if (_this.config.url == null || _this.config.list == null && args.listName === undefined) {
                reject("L'url ou la liste SP n'est pas défini.");
            }

            if (!$.isEmptyObject(_this.user.data)) {
                console.log("doublon !!");
                resolve(_this.user.data);
                return _this.user.data;

            }
            $.ajax({
                url: `${_this.config.url}/_api/web/lists/getbytitle('${listname}')/items?${_this._convertFiltersLine(args)}`,
                method: "GET",
                async: args.async,
                headers: {
                    "Accept": "application/json; odata=verbose"
                },
                dataType: 'json',
                success: function (msg) {
                    _this.user.data = msg;
                    resolve(_this.user.data);
                },
                error: function (msg) {
                    reject("Error :" + msg);
                }
            });
        })
    }

    fetch(args) {
        return this.fetchAll(Object.assign(args, {
            top: 1
        }));
    }

    clearData() {
        this.user.data = null;
        return true;
    }

    sendList(data) {
        var _this = this;

        var item = {
            "__metadata": {
                type: this.GetItemTypeForListName(this.getList())
            }
        };
        item = Object.assign(item, data);

        console.log("item", item)

        return new Promise(async (resolve, reject) => {

            var digest = false;
            await _this.getDigest().then((_digest) => { digest = _digest; })
    
            if(!digest) { reject("Error getting digest autorization."); }

            $.ajax({
                url: `${_this.config.url}/_api/web/lists/getbytitle('${this.getList()}')/items`,
                type: "POST",
                contentType: "application/json;odata=verbose",
                data: JSON.stringify(item),
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": digest
                },
                success: (data) => { resolve(data); },
                error: (msg) => { reject(`Erreur lors de l'envoi de donnée: ${msg}`) }
            })
        });
    }

    /**
     * 
     * data {
     *  filename,
     *  content,
     *  type,
     *  overwrite = false
     * }
     */
    uploadContent(data) {
        var _this = this;
        var listname = this.getList();
        
        if(data.overwrite === undefined) { data.overwrite = false; }
        if(data.listName) { listname = data.listName; }

        return new Promise(async (resolve, reject) => {

            var digest = false;

            await _this.getDigest().then((_digest) => { digest = _digest; })
    
            if(!digest) { reject("Error getting digest autorization."); }

            $.ajax({
                url: `${_this.config.url}/_api/web/GetFolderByServerRelativeUrl('${listname}')/Files/Add(url='${data.filename}', overwrite=${data.overwrite})`,
                type: "POST",
                contentType : "application/json;odata=verbose",
                data: data.content,
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-Type": data.type,
                    "X-RequestDigest": digest
                },
                success: function(req) {
                    resolve(req);
                    return req;
                },
                error: function(err) {
                    reject("Erreur "+err);
                }
            });
        })
       
    }

    getCurrentUser(async = false) {
        var _this = this;
        return new Promise((resolve, reject) => {
            if(_this.user.currentuser === null) {
                $.ajax({
                    async: async,
                    url: `${_this.config.url}/_api/web/currentuser`,
                    headers: { "Accept": "application/json; odata=verbose" },
                    dataType: 'json',
                    success: (msg) => { 
                        _this.user.currentuser = msg;
                        resolve(msg);
                        return msg;
                     },
                    error: (msg) => { reject(msg); }
                })
            }else {
                resolve(_this.user.currentuser);
            }
        })
    }

    getFileInFolder(name) {
        
        return new Promise(async (resolve, reject) => {
            var _this = this;
            var digest = false;

            await _this.getDigest().then((_digest) => { digest = _digest; })

            if(!digest) { reject("Error getting digest autorization."); }
            
            $.ajax({
                async: false,
                url: `${this.config.url}/_api/web/GetFolderByServerRelativeUrl('${this.config.list}')/Files('${name}')/$value`,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose",
                "Authorization" : "Bearer "+ digest },
                success: function(msg) {
                    resolve(msg);
                },
                error: function(err) {
                    reject(err);
                }
            })
        })
    }

}