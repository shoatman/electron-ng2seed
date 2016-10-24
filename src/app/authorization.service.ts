import { Injectable } from '@angular/core';
import {remote} from 'electron';
import {AuthenticationContextConfig, AuthenticationContext, Logger, LogLevel} from './adal';


/*
Notes:

- Need to detect when the user closes the popup rather than this code...

*/

interface CallBackFunc {
	(msg:string, token:any) : void;
}

@Injectable()
export class AuthorizationService {

	private authContext: any;
	private authWindow: Electron.BrowserWindow;

	public constructor() {

		var logger: Logger = {
			level : LogLevel.VERBOSE,
			log : function(msg:string){console.log(msg);}
		}	

		var authConfig : AuthenticationContextConfig = {
			instance: 'https://login.microsoftonline.com/',
      		tenant: 'common', //COMMON OR YOUR TENANT ID
      		clientId: '188743b4-4d95-4971-9d7b-1f2c105c9bda', //REPLACE WITH YOUR CLIENT ID
      		redirectUri: 'http://localhost', //REPLACE WITH YOUR REDIRECT URL
      		displayCall: this.electronSignIn.bind(this),
      		//callback: this.userSignedIn.bind(this),
      		popUp: false,
      		extraQueryParameter: 'prompt=login',
      		storage: window.localStorage,
      		crypto: window.crypto,
      		logger: logger
		};

		this.authContext = new AuthenticationContext( authConfig)
	}


	private createAuthWindow(){
		var webPreferences: Electron.WebPreferences = {};
		var options: Electron.BrowserWindowOptions = {};
		options.width = 800;
		options.height = 600;
		options.show = false;
		webPreferences.nodeIntegration = false;
		options.webPreferences = webPreferences;

		console.log(options);

		this.authWindow = new remote.BrowserWindow(options);


		var handleRedirect = function(event: Electron.Event, oldURL:string, newURL:string, isMainFrame:boolean, httpResponseCode:number, requestMethod:string, referrer:string, headers:Electron.Headers){
			this.handleCallBack(newURL);
			console.log('redirect detected');
		};

		var handleNavigation = function(event: Electron.Event, url:string, that:any){
			this.handleCallBack(url);
			console.log('navigation detected');
		};

		var handleRedirectBound = handleRedirect.bind(this);
		var handleNavigationBound = handleNavigation.bind(this);
		
		this.authWindow.webContents.on("did-get-redirect-request", handleRedirectBound);
		this.authWindow.webContents.on("will-navigate", handleNavigationBound);

	}

	private handleCallBack(url:string){
		if (url.indexOf(this.authContext.config.redirectUri) != -1) {
			console.log('got it');
			//this.authContext.handleWindowCallback(url.substr(url.indexOf('#'))); //This doesn't work due to checkfor popup=true...
			var hash = url.substr(url.indexOf('#'));
			var requestInfo = this.authContext.getRequestInfo(hash);
            this.authContext.saveTokenFromHash(requestInfo);
            this.authContext.callback("", requestInfo.parameters[this.authContext.CONSTANTS.ID_TOKEN]);
			this.authWindow.hide();
		}
	}

	private electronSignIn(url:string){
		console.log(url);

		this.createAuthWindow();
		this.authWindow.loadURL(url);
		this.authWindow.show();
	}

	public signIn(callback: CallBackFunc){
		this.authContext.callback = callback;
		this.authContext.login();
	}

	public acquireToken(callback: CallBackFunc, resource: string){
		this.authContext.callback = callback;
		this.authContext.acquireToken(resource);
	}

	public getCachedUser(){
		return this.authContext.getCachedUser();
	}


}