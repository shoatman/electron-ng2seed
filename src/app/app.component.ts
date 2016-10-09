import { Component } from '@angular/core';
import '../../public/css/styles.css';
import {AuthorizationService} from './authorization.service.ts';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
  providers: [AuthorizationService]
})
export class AppComponent { 

	constructor(private authorizationService: AuthorizationService){}

	authorize() {
		this.authorizationService.signIn();
	}

}
