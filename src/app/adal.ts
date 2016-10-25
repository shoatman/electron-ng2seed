//----------------------------------------------------------------------
// AdalJS v1.0.12 (Typescript)
//
//----------------------------------------------------------------------

    enum RequestType {
        Login = 1,
        RenewToken,
        Unknown
    };

    export enum LogLevel {
        ERROR = 0,
        WARN,
        INFO,
        VERBOSE
    }

    export const CONSTANTS:any = {
        ACCESS_TOKEN: 'access_token',
        EXPIRES_IN: 'expires_in',
        ID_TOKEN: 'id_token',
        ERROR_DESCRIPTION: 'error_description',
        SESSION_STATE: 'session_state',
        STORAGE: {
            TOKEN_KEYS: 'adal.token.keys',
            ACCESS_TOKEN_KEY: 'adal.access.token.key',
            EXPIRATION_KEY: 'adal.expiration.key',
            STATE_LOGIN: 'adal.state.login',
            STATE_RENEW: 'adal.state.renew',
            NONCE_IDTOKEN: 'adal.nonce.idtoken',
            SESSION_STATE: 'adal.session.state',
            USERNAME: 'adal.username',
            IDTOKEN: 'adal.idtoken',
            ERROR: 'adal.error',
            ERROR_DESCRIPTION: 'adal.error.description',
            LOGIN_REQUEST: 'adal.login.request',
            LOGIN_ERROR: 'adal.login.error',
            RENEW_STATUS: 'adal.token.renew.status'
        },
        RESOURCE_DELIMETER: '|',
        LOADFRAME_TIMEOUT: '6000',
        TOKEN_RENEW_STATUS_CANCELED: 'Canceled',
        TOKEN_RENEW_STATUS_COMPLETED: 'Completed',
        TOKEN_RENEW_STATUS_IN_PROGRESS: 'In Progress',
        LEVEL_STRING_MAP: {
            0: 'ERROR:',
            1: 'WARNING:',
            2: 'INFO:',
            3: 'VERBOSE:'
        },
        POPUP_WIDTH: 483,
        POPUP_HEIGHT: 600
    };

    export interface AuthenticationContextConfig {
        tenant: string;
        clientId: string,
        redirectUri: string,
        instance?: string,
        //endpoints: TBD,
        extraQueryParameter? : string;
        displayCall? : DisplayCallCallBackFunc;
        callback? : TokenReturnCallBackFunc;
        popUp?: boolean;
        isAngular?: boolean;
        logger: Logger;
        storage: Storage;
        crypto: Crypto;
        postLogoutRedirectUri?:string;
        expireOffsetSeconds?:number;
        loginResource?:string;        
    }

    export interface DisplayCallCallBackFunc {
        (url: string):void
    }

    export interface TokenReturnCallBackFunc {
        (err: string, token: string): void
    }

    export interface UserReturnCallBackFunc {
        (err:string, user:any):void
    }

    export interface LogFunc {
        (msg: string): void;
    }

    export interface Logger {
        level: LogLevel,
        log: LogFunc
    }


    /**
     * Config information
     * @public
     * @class Config
     * @property {tenant}          Your target tenant
     * @property {clientId}        Identifier assigned to your app by Azure Active Directory
     * @property {redirectUri}     Endpoint at which you expect to receive tokens
     * @property {instance}        Azure Active Directory Instance(default:https://login.microsoftonline.com/)
     * @property {endpoints}       Collection of {Endpoint-ResourceId} used for autmatically attaching tokens in webApi calls
     */
    export class AuthenticationContext {
      

        private _user: any = null;
        private _activeRenewals: any = {};
        private _loginInProgress: boolean = false;
        private _renewStates: any = [];
        private _callBackMappedToRenewStates: any = {};
        private _callBacksMappedToRenewStates: any = {};
        private _state: string;
        private _idTokenNonce: string;

        instance: string = 'https://login.microsoftonline.com/';
        config: AuthenticationContextConfig;
        callback: TokenReturnCallBackFunc;
        popUp: boolean = false;

        constructor (config: AuthenticationContextConfig){
            this.config = config;
            if(config.callback)
                this.callback = config.callback;
            if(config.popUp)
                this.popUp = config.popUp;
        }

        private log(level:LogLevel, message:string, error?:Error) {
            if (level <= this.config.logger.level) {
                var timestamp = new Date().toUTCString();
                var formattedMessage = '';
                formattedMessage = timestamp + ':' + this.libVersion() + '-' + CONSTANTS.LEVEL_STRING_MAP[level] + ' ' + message;

                this.config.logger.log(formattedMessage);
            }
        }

        private error(message:string, error:Error){
            this.log(LogLevel.ERROR, message, error);    
        }

        private warn(message:string){
            this.log(LogLevel.WARN, message);
        }

        private info(message:string){
            this.log(LogLevel.INFO, message);
        }

        private verbose(message:string){
            this.log(LogLevel.VERBOSE, message);
        }

        private cloneConfig(obj:any) {
            if (null === obj || 'object' !== typeof obj) {
                return obj;
            }

            var copy = {};
            for (var attr in obj) {
                if (obj.hasOwnProperty(attr)) {
                    copy[attr] = obj[attr];
                }
            }
            return copy;
        };

        private addLibMetadata():string {
            // x-client-SKU
            // x-client-Ver
            return '&x-client-SKU=Js&x-client-Ver=' + this.libVersion();
        };

        

        private libVersion():string {
            return '1.0.12';
        };

        private saveItem(key:string, obj:any) : boolean {
            this.config.storage.setItem(key, obj);
            return true;
        };

        private getItem (key:string): any {
           return this.config.storage.getItem(key);
        };

        /* jshint ignore:start */
        private generateGuid (): string {
            // RFC4122: The version 4 UUID is meant for generating UUIDs from truly-random or
            // pseudo-random numbers.
            // The algorithm is as follows:
            //     Set the two most significant bits (bits 6 and 7) of the
            //        clock_seq_hi_and_reserved to zero and one, respectively.
            //     Set the four most significant bits (bits 12 through 15) of the
            //        time_hi_and_version field to the 4-bit version number from
            //        Section 4.1.3. Version4
            //     Set all the other bits to randomly (or pseudo-randomly) chosen
            //     values.
            // UUID                   = time-low "-" time-mid "-"time-high-and-version "-"clock-seq-reserved and low(2hexOctet)"-" node
            // time-low               = 4hexOctet
            // time-mid               = 2hexOctet
            // time-high-and-version  = 2hexOctet
            // clock-seq-and-reserved = hexOctet:
            // clock-seq-low          = hexOctet
            // node                   = 6hexOctet
            // Format: xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx
            // y could be 1000, 1001, 1010, 1011 since most significant two bits needs to be 10
            // y values are 8, 9, A, B
            
            if (this.config.crypto && this.config.crypto.getRandomValues) {
                var buffer = new Uint8Array(16);
                this.config.crypto.getRandomValues(buffer);
                //buffer[6] and buffer[7] represents the time_hi_and_version field. We will set the four most significant bits (4 through 7) of buffer[6] to represent decimal number 4 (UUID version number).
                buffer[6] |= 0x40; //buffer[6] | 01000000 will set the 6 bit to 1.
                buffer[6] &= 0x4f; //buffer[6] & 01001111 will set the 4, 5, and 7 bit to 0 such that bits 4-7 == 0100 = "4".
                //buffer[8] represents the clock_seq_hi_and_reserved field. We will set the two most significant bits (6 and 7) of the clock_seq_hi_and_reserved to zero and one, respectively.
                buffer[8] |= 0x80; //buffer[8] | 10000000 will set the 7 bit to 1.
                buffer[8] &= 0xbf; //buffer[8] & 10111111 will set the 6 bit to 0.
                return this.decimalToHex(buffer[0]) + this.decimalToHex(buffer[1]) + this.decimalToHex(buffer[2]) + this.decimalToHex(buffer[3]) + '-' + this.decimalToHex(buffer[4]) + this.decimalToHex(buffer[5]) + '-' + this.decimalToHex(buffer[6]) + this.decimalToHex(buffer[7]) + '-' +
                 this.decimalToHex(buffer[8]) + this.decimalToHex(buffer[9]) + '-' + this.decimalToHex(buffer[10]) + this.decimalToHex(buffer[11]) + this.decimalToHex(buffer[12]) + this.decimalToHex(buffer[13]) + this.decimalToHex(buffer[14]) + this.decimalToHex(buffer[15]);
            }
            else {
                var guidHolder = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
                var hex = '0123456789abcdef';
                var r = 0;
                var guidResponse = "";
                for (var i = 0; i < 36; i++) {
                    if (guidHolder[i] !== '-' && guidHolder[i] !== '4') {
                        // each x and y needs to be random
                        r = Math.random() * 16 | 0;
                    }
                    if (guidHolder[i] === 'x') {
                        guidResponse += hex[r];
                    } else if (guidHolder[i] === 'y') {
                        // clock-seq-and-reserved first hex is filtered and remaining hex values are random
                        r &= 0x3; // bit and with 0011 to set pos 2 to zero ?0??
                        r |= 0x8; // set pos 3 to 1 as 1???
                        guidResponse += hex[r];
                    } else {
                        guidResponse += guidHolder[i];
                    }
                }
                return guidResponse;
            }
        };
        /* jshint ignore:end */

        private expiresIn(expires:string): number {
            return this.now() + parseInt(expires, 10);
        };

        private now():number {
            return Math.round(new Date().getTime() / 1000.0);
        };

        private isEmpty(str:string) {
            return (typeof str === 'undefined' || !str || 0 === str.length);
        };

        private addAdalFrame(iframeId:string):any {
            if (typeof iframeId === 'undefined') {
                return;
            }

            this.info('Add adal frame to document:' + iframeId);
            var adalFrame: any = document.getElementById(iframeId);

            if (!adalFrame) {
                if (document.createElement && document.documentElement &&
                    (window['opera'] || window.navigator.userAgent.indexOf('MSIE 5.0') === -1)) {
                    var ifr = document.createElement('iframe');
                    ifr.setAttribute('id', iframeId);
                    ifr.style.visibility = 'hidden';
                    ifr.style.position = 'absolute';
                    ifr.style.width = ifr.style.height = ifr['borderWidth'] = '0px';

                    adalFrame = document.getElementsByTagName('body')[0].appendChild(ifr);
                }
                else if (document.body && document.body.insertAdjacentHTML) {
                    document.body.insertAdjacentHTML('beforeEnd', '<iframe name="' + iframeId + '" id="' + iframeId + '" style="display:none"></iframe>');
                }
                if (window.frames && window.frames[iframeId]) {
                    adalFrame = window.frames[iframeId];
                }
            }

            return adalFrame;
        };

        private convertUrlSafeToRegularBase64EncodedString(str:string):string {
            return str.replace('-', '+').replace('_', '/');
        };

        private serialize(responseType:string, obj:any, resource:string) {
            var str:any = [];
            if (obj !== null) {
                str.push('?response_type=' + responseType);
                str.push('client_id=' + encodeURIComponent(obj.clientId));
                if (resource) {
                    str.push('resource=' + encodeURIComponent(resource));
                }

                str.push('redirect_uri=' + encodeURIComponent(obj.redirectUri));
                str.push('state=' + encodeURIComponent(this._state));

                if (obj.hasOwnProperty('slice')) {
                    str.push('slice=' + encodeURIComponent(obj.slice));
                }

                if (obj.hasOwnProperty('extraQueryParameter')) {
                    str.push(obj.extraQueryParameter);
                }

                var correlationId = obj.correlationId ? obj.correlationId : this.generateGuid();
                str.push('client-request-id=' + encodeURIComponent(correlationId));
            }

            return str.join('&');
        };

        private deserialize(query:string):any {
            var match: any;
            var pl:RegExp = /\+/g;  // Regex for replacing addition symbol with a space
            var search:RegExp = /([^&=]+)=([^&]*)/g;
            var decode = function (s:string) {
                    return decodeURIComponent(s.replace(pl, ' '));
                };
            var obj = {};
            match = search.exec(query);
            while (match) {
                obj[decode(match[1])] = decode(match[2]);
                match = search.exec(query);
            }

            this.verbose("Deserialized Object from hash");
            console.log(obj);
            return obj;
        };

        private decimalToHex(number:number):string {
            var hex = number.toString(16);
            while (hex.length < 2) {
                hex = '0' + hex;
            }
            return hex;
        }

        private getNavigateUrl(responseType:string, resource:string) {
            var tenant = 'common';
            if (this.config.tenant) {
                tenant = this.config.tenant;
            }

            var urlNavigate = this.instance + tenant + '/oauth2/authorize' + this.serialize(responseType, this.config, resource) + this.addLibMetadata();
            this.info('Navigate url:' + urlNavigate);
            return urlNavigate;
        };

        private extractIdToken(encodedIdToken:any) {
            // id token will be decoded to get the username
            var decodedToken = this.decodeJwt(encodedIdToken);
            if (!decodedToken) {
                return null;
            }

            try {
                var base64IdToken = decodedToken.JWSPayload;
                var base64Decoded = this.base64DecodeStringUrlSafe(base64IdToken);
                if (!base64Decoded) {
                    this.info('The returned id_token could not be base64 url safe decoded.');
                    return null;
                }

                // ECMA script has JSON built-in support
                return JSON.parse(base64Decoded);
            } catch (err) {
                this.error('The returned id_token could not be decoded', err);
            }

            return null;
        };

        private base64DecodeStringUrlSafe(base64IdToken:string) {
            // html5 should support atob function for decoding
            base64IdToken = base64IdToken.replace(/-/g, '+').replace(/_/g, '/');
            if (window.atob) {
                return decodeURIComponent(encodeURI(window.atob(base64IdToken))); // jshint ignore:line
            }
            else {
                return decodeURIComponent(encodeURI(this.decode(base64IdToken)));
            }
        };

        //Take https://cdnjs.cloudflare.com/ajax/libs/Base64/0.3.0/base64.js and https://en.wikipedia.org/wiki/Base64 as reference. 
        private decode(base64IdToken:string):string {
            var codes = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=';
            base64IdToken = String(base64IdToken).replace(/=+$/, '');

            var length = base64IdToken.length;
            if (length % 4 === 1) {
                throw new Error('The token to be decoded is not correctly encoded.');
            }

            var h1:any;
            var h2:any; 
            var h3:any; 
            var h4:any;
            var bits:any; 
            var c1:any;
            var c2:any;
            var c3:any; 
            var decoded:string = '';
            for (var i = 0; i < length; i += 4) {
                //Every 4 base64 encoded character will be converted to 3 byte string, which is 24 bits
                // then 6 bits per base64 encoded character
                h1 = codes.indexOf(base64IdToken.charAt(i));
                h2 = codes.indexOf(base64IdToken.charAt(i + 1));
                h3 = codes.indexOf(base64IdToken.charAt(i + 2));
                h4 = codes.indexOf(base64IdToken.charAt(i + 3));

                // For padding, if last two are '='
                if (i + 2 === length - 1) {
                    bits = h1 << 18 | h2 << 12 | h3 << 6;
                    c1 = bits >> 16 & 255;
                    c2 = bits >> 8 & 255;
                    decoded += String.fromCharCode(c1, c2);
                    break;
                }
                    // if last one is '='
                else if (i + 1 === length - 1) {
                    bits = h1 << 18 | h2 << 12
                    c1 = bits >> 16 & 255;
                    decoded += String.fromCharCode(c1);
                    break;
                }

                bits = h1 << 18 | h2 << 12 | h3 << 6 | h4;

                // then convert to 3 byte chars
                c1 = bits >> 16 & 255;
                c2 = bits >> 8 & 255;
                c3 = bits & 255;

                decoded += String.fromCharCode(c1, c2, c3);
            }

            return decoded;
        };

        // Adal.node js crack function
        private decodeJwt(jwtToken:any) {
            if (this.isEmpty(jwtToken)) {
                return null;
            };

            var idTokenPartsRegex = /^([^\.\s]*)\.([^\.\s]+)\.([^\.\s]*)$/;

            var matches = idTokenPartsRegex.exec(jwtToken);
            if (!matches || matches.length < 4) {
                this.warn('The returned id_token is not parseable.');
                return null;
            }

            var crackedToken = {
                header: matches[1],
                JWSPayload: matches[2],
                JWSSig: matches[3]
            };

            return crackedToken;
        };

        private hasResource(key:string):any {
            var keys = this.getItem(CONSTANTS.STORAGE.TOKEN_KEYS);
            return keys && !this.isEmpty(keys) && (keys.indexOf(key + CONSTANTS.RESOURCE_DELIMETER) > -1);
        };

        /**
         * Gets token for the specified resource from local storage cache
         * @param {string}   resource A URI that identifies the resource for which the token is valid.
         * @returns {string} token if exists and not expired or null
         */
        getCachedToken(resource:string):any {
            if (!this.hasResource(resource)) {
                return null;
            }

            var token = this.getItem(CONSTANTS.STORAGE.ACCESS_TOKEN_KEY + resource);
            var expired = this.getItem(CONSTANTS.STORAGE.EXPIRATION_KEY + resource);

            // If expiration is within offset, it will force renew
            var offset = this.config.expireOffsetSeconds || 120;  //Hmm

            if (expired && (expired > this.now() + offset)) {
                return token;
            } else {
                this.saveItem(CONSTANTS.STORAGE.ACCESS_TOKEN_KEY + resource, '');
                this.saveItem(CONSTANTS.STORAGE.EXPIRATION_KEY + resource, 0);
                return null;
            }
        };
        /**
         * Retrieves and parse idToken from localstorage
         * @returns {User} user object
         */
        getCachedUser():any {
            if (this._user) {
                return this._user;
            }

            var idtoken = this.getItem(CONSTANTS.STORAGE.IDTOKEN);
            this._user = this.createUser(idtoken);
            return this._user;
        };

        registerCallback(expectedState:string, resource:string, callback:TokenReturnCallBackFunc) {
            this._activeRenewals[resource] = expectedState;
            if (!this._callBacksMappedToRenewStates[expectedState]) {
                this._callBacksMappedToRenewStates[expectedState] = [];
            }
            var self = this;
            this._callBacksMappedToRenewStates[expectedState].push(callback);
            if (!this._callBackMappedToRenewStates[expectedState]) {
                this._callBackMappedToRenewStates[expectedState] = function (message:string, token:any) {
                    for (var i = 0; i < self._callBacksMappedToRenewStates[expectedState].length; ++i) {
                        try {
                            self._callBacksMappedToRenewStates[expectedState][i](message, token);
                        }
                        catch (error) {
                            self.warn(error);
                        }
                    }
                    self._activeRenewals[resource] = null;
                    self._callBacksMappedToRenewStates[expectedState] = null;
                    self._callBackMappedToRenewStates[expectedState] = null;
                };
            }
        };

        private renewToken(resource:string, callback:TokenReturnCallBackFunc) {
            // use iframe to try refresh token
            // use given resource to create new authz url
            this.info('renewToken is called for resource:' + resource);
            var frameHandle = this.addAdalFrame('adalRenewFrame' + resource);
            var expectedState = this.generateGuid() + '|' + resource;
            this._state = expectedState;
            // renew happens in iframe, so it keeps javascript context
            this._renewStates.push(expectedState);

            this.verbose('Renew token Expected state: ' + expectedState);
            var urlNavigate = this.getNavigateUrl('token', resource) + '&prompt=none';
            urlNavigate = this.addHintParameters(urlNavigate);

            this.registerCallback(expectedState, resource, callback);
            this.verbose('Navigate to:' + urlNavigate);
            frameHandle.src = 'about:blank';
            this.loadFrameTimeout(urlNavigate, 'adalRenewFrame' + resource, resource);

        };

        private renewIdToken(callback:TokenReturnCallBackFunc) {
            // use iframe to try refresh token
            this.info('renewIdToken is called');
            var frameHandle = this.addAdalFrame('adalIdTokenFrame');
            var expectedState = this.generateGuid() + '|' + this.config.clientId;
            this._idTokenNonce = this.generateGuid();
            this.saveItem(CONSTANTS.STORAGE.NONCE_IDTOKEN, this._idTokenNonce);
            this._state = expectedState;
            // renew happens in iframe, so it keeps javascript context
            this._renewStates.push(expectedState);

            this.verbose('Renew Idtoken Expected state: ' + expectedState);
            var urlNavigate = this.getNavigateUrl('id_token', null) + '&prompt=none';
            urlNavigate = this.addHintParameters(urlNavigate);

            urlNavigate += '&nonce=' + encodeURIComponent(this._idTokenNonce);
            this.registerCallback(expectedState, this.config.clientId, callback);
            this._idTokenNonce = null;
            this.verbose('Navigate to:' + urlNavigate);
            frameHandle.src = 'about:blank';
            this.loadFrameTimeout(urlNavigate, 'adalIdTokenFrame', this.config.clientId);
        };

        private urlContainsQueryStringParameter(name:string, url:string) {
            // regex to detect pattern of a ? or & followed by the name parameter and an equals character
            var regex = new RegExp("[\\?&]" + name + "=");
            return regex.test(url);
        }

        // Calling _loadFrame but with a timeout to signal failure in loadframeStatus. Callbacks are left
        // registered when network errors occur and subsequent token requests for same resource are registered to the pending request
        private loadFrameTimeout(urlNavigation:string, frameName:string, resource:string) {
            //set iframe session to pending
            this.verbose('Set loading state to pending for: ' + resource);
            this.saveItem(CONSTANTS.STORAGE.RENEW_STATUS + resource, CONSTANTS.TOKEN_RENEW_STATUS_IN_PROGRESS);
            this.loadFrame(urlNavigation, frameName);
            var self = this;
            setTimeout(function () {
                if (self.getItem(CONSTANTS.STORAGE.RENEW_STATUS + resource) === CONSTANTS.TOKEN_RENEW_STATUS_IN_PROGRESS) {
                    // fail the iframe session if it's in pending state
                    self.verbose('Loading frame has timed out after: ' + (CONSTANTS.LOADFRAME_TIMEOUT / 1000).toString() + ' seconds for resource ' + resource);
                    var expectedState = self._activeRenewals[resource];
                    if (expectedState && self._callBackMappedToRenewStates[expectedState]) {
                        self._callBackMappedToRenewStates[expectedState]('Token renewal operation failed due to timeout', null);
                    }

                    self.saveItem(CONSTANTS.STORAGE.RENEW_STATUS + resource, CONSTANTS.TOKEN_RENEW_STATUS_CANCELED);
                }
            }, CONSTANTS.LOADFRAME_TIMEOUT);
        }

        private loadFrame(urlNavigate:string, frameName:string) {
            // This trick overcomes iframe navigation in IE
            // IE does not load the page consistently in iframe
            var self = this;
            self.info('LoadFrame: ' + frameName);
            var frameCheck = frameName;
            setTimeout(function () {
                var frameHandle = self.addAdalFrame(frameCheck);
                if (frameHandle.src === '' || frameHandle.src === 'about:blank') {
                    frameHandle.src = urlNavigate;
                    self.loadFrame(urlNavigate, frameCheck);
                }
            }, 500);
        };

        /**
         * Acquire token from cache if not expired and available. Acquires token from iframe if expired.
         * @param {string}   resource  ResourceUri identifying the target resource
         * @param {requestCallback} callback
         */
        acquireToken(resource:string, callback:TokenReturnCallBackFunc) {
            
            this.verbose("acquireToken Enter");

            if (this.isEmpty(resource)) {
                this.warn('resource is required');
                callback('resource is required', null);
                return;
            }

            var token = this.getCachedToken(resource);
            if (token) {
                this.info('Token is already in cache for resource:' + resource);
                callback(null, token);
                return;
            }

            if (!this._user) {
                this.warn('User login is required');
                callback('User login is required', null);
                return;
            }

            // refresh attept with iframe
            //Already renewing for this resource, callback when we get the token.
            if (this._activeRenewals[resource]) {
                //Active renewals contains the state for each renewal.
                this.registerCallback(this._activeRenewals[resource], resource, callback);
            }
            else {
                if (resource === this.config.clientId) {
                    // App uses idtoken to send to api endpoints
                    // Default resource is tracked as clientid to store this token
                    this.verbose('renewing idtoken');
                    this.renewIdToken(callback);
                } else {
                    this.verbose('renewing token');
                    this.renewToken(resource, callback);
                }
            }

            this.verbose("acquireToken exit");
        };

        /**
         * Redirect the Browser to Azure AD Authorization endpoint
         * @param {string}   urlNavigate The authorization request url
         */
        promptUser(urlNavigate:string) {
            if (urlNavigate) {
                this.info('Navigate to:' + urlNavigate);
                window.location.replace(urlNavigate);
            } else {
                this.info('Navigate url is empty');
            }
        };

        /**
         * Clear cache items.
         */
        clearCache():void {
            this.saveItem(CONSTANTS.STORAGE.ACCESS_TOKEN_KEY, '');
            this.saveItem(CONSTANTS.STORAGE.EXPIRATION_KEY, 0);
            this.saveItem(CONSTANTS.STORAGE.SESSION_STATE, '');
            this.saveItem(CONSTANTS.STORAGE.STATE_LOGIN, '');
            this._renewStates = [];
            this.saveItem(CONSTANTS.STORAGE.USERNAME, '');
            this.saveItem(CONSTANTS.STORAGE.IDTOKEN, '');
            this.saveItem(CONSTANTS.STORAGE.ERROR, '');
            this.saveItem(CONSTANTS.STORAGE.ERROR_DESCRIPTION, '');
            var keys = this.getItem(CONSTANTS.STORAGE.TOKEN_KEYS);

            if (!this.isEmpty(keys)) {
                keys = keys.split(CONSTANTS.RESOURCE_DELIMETER);
                for (var i = 0; i < keys.length; i++) {
                    this.saveItem(CONSTANTS.STORAGE.ACCESS_TOKEN_KEY + keys[i], '');
                    this.saveItem(CONSTANTS.STORAGE.EXPIRATION_KEY + keys[i], 0);
                }
            }
            this.saveItem(CONSTANTS.STORAGE.TOKEN_KEYS, '');
        };

        /**
         * Clear cache items for a resource.
         */
        clearCacheForResource(resource:string):void {
            this.saveItem(CONSTANTS.STORAGE.STATE_RENEW, '');
            this.saveItem(CONSTANTS.STORAGE.ERROR, '');
            this.saveItem(CONSTANTS.STORAGE.ERROR_DESCRIPTION, '');
            if (this.hasResource(resource)) {
                this.saveItem(CONSTANTS.STORAGE.ACCESS_TOKEN_KEY + resource, '');
                this.saveItem(CONSTANTS.STORAGE.EXPIRATION_KEY + resource, 0);
            }
        };

        /**
         * Logout user will redirect page to logout endpoint.
         * After logout, it will redirect to post_logout page if provided.
         */
        logOut() {
            this.clearCache();
            var tenant = 'common';
            var logout = '';
            this._user = null;
            if (this.config.tenant) {
                tenant = this.config.tenant;
            }

            if (this.config.postLogoutRedirectUri) {
                logout = 'post_logout_redirect_uri=' + encodeURIComponent(this.config.postLogoutRedirectUri);
            }

            var urlNavigate = this.instance + tenant + '/oauth2/logout?' + logout;
            this.info('Logout navigate to: ' + urlNavigate);
            this.promptUser(urlNavigate);
        };

        

        /**
         * This callback is displayed as part of the Requester class.
         * @callback requestCallback
         * @param {string} error
         * @param {User} user
         */

        /**
         * Gets a user profile
         * @param {requestCallback} cb - The callback that handles the response.
         */
        getUser(callback:UserReturnCallBackFunc) {
            // IDToken is first call
            if (typeof callback !== 'function') {
                throw new Error('callback is not a function');
            }

            // user in memory
            if (this._user) {
                callback(null, this._user);
                return;
            }

            // frame is used to get idtoken
            var idtoken = this.getItem(CONSTANTS.STORAGE.IDTOKEN);
            if (!this.isEmpty(idtoken)) {
                this.info('User exists in cache: ');
                this._user = this.createUser(idtoken);
                callback(null, this._user);
            } else {
                this.warn('User information is not available');
                callback('User information is not available', null);
            }
        };

        private addHintParameters(urlNavigate:string) {
            // include hint params only if upn is present
            if (this._user && this._user.profile && this._user.profile.hasOwnProperty('upn')) {

                // add login_hint
                urlNavigate += '&login_hint=' + encodeURIComponent(this._user.profile.upn);

                // don't add domain_hint twice if user provided it in the extraQueryParameter value
                if (!this.urlContainsQueryStringParameter("domain_hint", urlNavigate) && this._user.profile.upn.indexOf('@') > -1) {
                    var parts = this._user.profile.upn.split('@');
                    // local part can include @ in quotes. Sending last part handles that.
                    urlNavigate += '&domain_hint=' + encodeURIComponent(parts[parts.length - 1]);
                }
            }

            return urlNavigate;
        }

        private createUser(idToken:string) {
            var user:any = null;
            var parsedJson:any = this.extractIdToken(idToken);
            if (parsedJson && parsedJson.hasOwnProperty('aud')) {
                if (parsedJson.aud.toLowerCase() === this.config.clientId.toLowerCase()) {

                    user = {
                        userName: '',
                        profile: parsedJson
                    };

                    if (parsedJson.hasOwnProperty('upn')) {
                        user.userName = parsedJson.upn;
                    } else if (parsedJson.hasOwnProperty('email')) {
                        user.userName = parsedJson.email;
                    }
                } else {
                    this.warn('IdToken has invalid aud field');
                }

            }

            return user;
        };

        private getHash(hash:string):string {
            if (hash.indexOf('#/') > -1) {
                hash = hash.substring(hash.indexOf('#/') + 2);
            } else if (hash.indexOf('#') > -1) {
                hash = hash.substring(1);
            }

            return hash;
        };

        /**
         * Checks if hash contains access token or id token or error_description
         * @param {string} hash  -  Hash passed from redirect page
         * @returns {Boolean}
         */
        isCallback(hash:string) {
            hash = this.getHash(hash);
            var parameters = this.deserialize(hash);
            return (
                parameters.hasOwnProperty(CONSTANTS.ERROR_DESCRIPTION) ||
                parameters.hasOwnProperty(CONSTANTS.ACCESS_TOKEN) ||
                parameters.hasOwnProperty(CONSTANTS.ID_TOKEN)
            );
        };

        /**
         * Gets login error
         * @returns {string} error message related to login
         */
        getLoginError():string {
            return this.getItem(CONSTANTS.STORAGE.LOGIN_ERROR);
        };

        /**
         * Gets requestInfo from given hash.
         * @returns {string} error message related to login
         */
        getRequestInfo(hash:string) {
            this.verbose("getRequestInfo enter");
            hash = this.getHash(hash);
            var parameters = this.deserialize(hash);
            var requestInfo = {
                valid: false,
                parameters: {},
                stateMatch: false,
                stateResponse: '',
                requestType: RequestType.Unknown
            };


            if (parameters) {
                this.verbose("getRequestInfo: Parameters found");
                this.verbose(parameters.toString());
                requestInfo.parameters = parameters;
                if (parameters.hasOwnProperty(CONSTANTS.ERROR_DESCRIPTION) ||
                    parameters.hasOwnProperty(CONSTANTS.ACCESS_TOKEN) ||
                    parameters.hasOwnProperty(CONSTANTS.ID_TOKEN)) {

                    requestInfo.valid = true;

                    // which call
                    var stateResponse = '';
                    if (parameters.hasOwnProperty('state')) {
                        this.verbose('State: ' + parameters.state);
                        stateResponse = parameters.state;
                    } else {
                        this.warn('No state returned');
                        return requestInfo;
                    }

                    requestInfo.stateResponse = stateResponse;

                    this.verbose(CONSTANTS.STORAGE.STATE_LOGIN);
                    this.verbose(this.getItem(CONSTANTS.STORAGE.STATE_LOGIN));

                    // async calls can fire iframe and login request at the same time if developer does not use the API as expected
                    // incoming callback needs to be looked up to find the request type
                    if (stateResponse === this.getItem(CONSTANTS.STORAGE.STATE_LOGIN)) {
                        this.verbose("stored state and returned state match");
                        requestInfo.requestType = RequestType.Login;
                        requestInfo.stateMatch = true;
                        return requestInfo;
                    }else{
                        this.warn("states do not match");
                    }

                    /*
                    NOTE: I don't think this applies in any situation where the ADAL should be used... but I could be wrong
                    // external api requests may have many renewtoken requests for different resource
                    if (!requestInfo.stateMatch && window.parent && window.parent.AuthenticationContext) {
                        var statesInParentContext = window.parent.AuthenticationContext()._renewStates;
                        for (var i = 0; i < statesInParentContext.length; i++) {
                            if (statesInParentContext[i] === requestInfo.stateResponse) {
                                requestInfo.requestType = this.REQUEST_TYPE.RENEW_TOKEN;
                                requestInfo.stateMatch = true;
                                break;
                            }
                        }
                    }
                    */
                }
            }

            return requestInfo;
        };

        private getResourceFromState(state:string):string {
            if (state) {
                var splitIndex = state.indexOf('|');
                if (splitIndex > -1 && splitIndex + 1 < state.length) {
                    return state.substring(splitIndex + 1);
                }
            }

            return '';
        };

        /**
         * Saves token from hash that is received from redirect.
         * @param {string} hash  -  Hash passed from redirect page
         * @returns {string} error message related to login
         */
        saveTokenFromHash(requestInfo:any) {
            this.info('State status:' + requestInfo.stateMatch + '; Request type:' + requestInfo.requestType);
            this.saveItem(CONSTANTS.STORAGE.ERROR, '');
            this.saveItem(CONSTANTS.STORAGE.ERROR_DESCRIPTION, '');

            var resource = this.getResourceFromState(requestInfo.stateResponse);

            // Record error
            if (requestInfo.parameters.hasOwnProperty(CONSTANTS.ERROR_DESCRIPTION)) {
                this.info('Error :' + requestInfo.parameters.error + '; Error description:' + requestInfo.parameters[CONSTANTS.ERROR_DESCRIPTION]);
                this.saveItem(CONSTANTS.STORAGE.ERROR, requestInfo.parameters.error);
                this.saveItem(CONSTANTS.STORAGE.ERROR_DESCRIPTION, requestInfo.parameters[CONSTANTS.ERROR_DESCRIPTION]);

                if (requestInfo.requestType === RequestType.Login) {
                    this._loginInProgress = false;
                    this.saveItem(CONSTANTS.STORAGE.LOGIN_ERROR, requestInfo.parameters.error_description);
                }
            } else {
                // It must verify the state from redirect
                if (requestInfo.stateMatch) {
                    // record tokens to storage if exists
                    this.info('State is right');
                    if (requestInfo.parameters.hasOwnProperty(CONSTANTS.SESSION_STATE)) {
                        this.saveItem(CONSTANTS.STORAGE.SESSION_STATE, requestInfo.parameters[CONSTANTS.SESSION_STATE]);
                    }

                    var keys:string;

                    if (requestInfo.parameters.hasOwnProperty(CONSTANTS.ACCESS_TOKEN)) {
                        this.info('Fragment has access token');

                        if (!this.hasResource(resource)) {
                            keys = this.getItem(CONSTANTS.STORAGE.TOKEN_KEYS) || '';
                            this.saveItem(CONSTANTS.STORAGE.TOKEN_KEYS, keys + resource + CONSTANTS.RESOURCE_DELIMETER);
                        }
                        // save token with related resource
                        this.saveItem(CONSTANTS.STORAGE.ACCESS_TOKEN_KEY + resource, requestInfo.parameters[CONSTANTS.ACCESS_TOKEN]);
                        this.saveItem(CONSTANTS.STORAGE.EXPIRATION_KEY + resource, this.expiresIn(requestInfo.parameters[CONSTANTS.EXPIRES_IN]));
                    }

                    if (requestInfo.parameters.hasOwnProperty(CONSTANTS.ID_TOKEN)) {
                        this.info('Fragment has id token');
                        this._loginInProgress = false;

                        this._user = this.createUser(requestInfo.parameters[CONSTANTS.ID_TOKEN]);

                        if (this._user && this._user.profile) {
                            if (this._user.profile.nonce !== this.getItem(CONSTANTS.STORAGE.NONCE_IDTOKEN)) {
                                this._user = null;
                                this.saveItem(CONSTANTS.STORAGE.LOGIN_ERROR, 'Nonce is not same as ' + this._idTokenNonce);
                            } else {
                                this.saveItem(CONSTANTS.STORAGE.IDTOKEN, requestInfo.parameters[CONSTANTS.ID_TOKEN]);

                                // Save idtoken as access token for app itself
                                resource = this.config.loginResource ? this.config.loginResource : this.config.clientId;

                                if (!this.hasResource(resource)) {
                                    keys = this.getItem(CONSTANTS.STORAGE.TOKEN_KEYS) || '';
                                    this.saveItem(CONSTANTS.STORAGE.TOKEN_KEYS, keys + resource + CONSTANTS.RESOURCE_DELIMETER);
                                }
                                this.saveItem(CONSTANTS.STORAGE.ACCESS_TOKEN_KEY + resource, requestInfo.parameters[CONSTANTS.ID_TOKEN]);
                                this.saveItem(CONSTANTS.STORAGE.EXPIRATION_KEY + resource, this._user.profile.exp);
                            }
                        }
                        else {
                            this.saveItem(CONSTANTS.STORAGE.ERROR, 'invalid id_token');
                            this.saveItem(CONSTANTS.STORAGE.ERROR_DESCRIPTION, 'Invalid id_token. id_token: ' + requestInfo.parameters[CONSTANTS.ID_TOKEN]);
                        }
                    }
                } else {
                    this.saveItem(CONSTANTS.STORAGE.ERROR, 'Invalid_state');
                    this.saveItem(CONSTANTS.STORAGE.ERROR_DESCRIPTION, 'Invalid_state. state: ' + requestInfo.stateResponse);
                }
            }
            this.saveItem(CONSTANTS.STORAGE.RENEW_STATUS + resource, CONSTANTS.TOKEN_RENEW_STATUS_COMPLETED);
        };

        login(){
            if (this._loginInProgress) {
                this.info("Login already in progress");
                return;
            }

            this._state = this.generateGuid();
            this._idTokenNonce = this.generateGuid();

            this.verbose('Expected state: ' + this._state + ' startPage:' + window.location);

            this.saveItem(CONSTANTS.STORAGE.LOGIN_REQUEST, window.location);
            this.saveItem(CONSTANTS.STORAGE.LOGIN_ERROR, '');
            this.saveItem(CONSTANTS.STORAGE.STATE_LOGIN, this._state);
            this.saveItem(CONSTANTS.STORAGE.NONCE_IDTOKEN, this._idTokenNonce);
            this.saveItem(CONSTANTS.STORAGE.ERROR, '');
            this.saveItem(CONSTANTS.STORAGE.ERROR_DESCRIPTION, '');

            var urlNavigate = this.getNavigateUrl('id_token', null) + '&nonce=' + encodeURIComponent(this._idTokenNonce);
            this._loginInProgress = true;

            if (this.config.displayCall) {
                // User defined way of handling the navigation
                this.config.displayCall(urlNavigate);
            } else {
                this.promptUser(urlNavigate);
            }
        }

    };
