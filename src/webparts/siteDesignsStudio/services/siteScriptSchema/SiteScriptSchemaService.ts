import { ISiteScriptAction } from '../../models/ISiteScript';
import Schema from '../../schema/schema';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';

export interface ISiteScriptSchemaService {
	getActionSchemaAsync(action: ISiteScriptAction): Promise<any>;
	getActionSchema(action: ISiteScriptAction): any;
	getSubActionSchema(parentAction: ISiteScriptAction, subAction: ISiteScriptAction): any;
	getAvailableActionsAsync(): Promise<string[]>;
	getAvailableActions(): string[];
  getAvailableSubActions(parentAction: ISiteScriptAction): string[];
  getAvailableSubActionsAsync(parentAction: ISiteScriptAction): Promise<string[]>;
}

export class SiteScriptSchemaService implements ISiteScriptSchemaService {
	private isInitialized: boolean = false;
	private availableActions: string[] = null;
	private availableSubActionByVerb: {} = null;
	private availableActionSchemas = null;
	private availableSubActionSchemasByVerb = null;

	constructor(serviceScope: ServiceScope) {}

	private ensureInitialized(): Promise<void> {
		return this._getSchema().then((schema) => {
			if (this.isInitialized) {
				return;
			}

			// Get available action schemas
			let actionsArraySchema = schema.properties.actions;

			if (!actionsArraySchema.type || actionsArraySchema.type != 'array') {
				throw new Error('Invalid Actions schema');
			}

			if (!actionsArraySchema.items || !actionsArraySchema.items.anyOf) {
				throw new Error('Invalid Actions schema');
			}

			let actionsArraySchemaItems = actionsArraySchema.items;

			// Get Main Actions schema
			let availableActionSchemasAsArray: any[] = actionsArraySchemaItems.anyOf.map((action) =>
				this._getElementSchema(schema, action)
			);
			this.availableActionSchemas = {};
			availableActionSchemasAsArray.forEach((a) => {
				// Keep the current action schema
				let actionVerb = this._getVerbFromActionSchema(a);
				this.availableActionSchemas[actionVerb] = a;

				// Check if the current action has subactions
				let subActionSchemas = this._getSubActionsSchemaFromParentActionSchema(schema, a);
				if (subActionSchemas) {
					// If yes, keep the sub actions schema and verbs

					// Keep the list of subactions verbs
					if (!this.availableSubActionByVerb) {
						this.availableSubActionByVerb = {};
					}
					this.availableSubActionByVerb[actionVerb] = subActionSchemas.map((sa) =>
						this._getVerbFromActionSchema(sa)
					);

					// Keep the list of subactions schemas
					if (!this.availableSubActionSchemasByVerb) {
						this.availableSubActionSchemasByVerb = {};
					}
					this.availableSubActionSchemasByVerb[actionVerb] = {};
					subActionSchemas.forEach((sas) => {
						let subActionVerb = this._getVerbFromActionSchema(sas);
						this.availableSubActionSchemasByVerb[actionVerb][subActionVerb] = sas;
					});
				}
			});
			this.availableActions = availableActionSchemasAsArray.map((a) => this._getVerbFromActionSchema(a));

			// // For each of the any Of, get the appropriate definition
			// let definitionRelativeReferences: any[] = actionsArraySchemaItems.anyOf.map((d) =>
			// 	d['$ref'].replace('#/definitions/', '')
			// );

			// this.availableActionSchemas = {};
			// this.availableActionSchemasByVerb = {};
			// this.availableSubActionSchemasByVerb = {};
			// this.availableActions = [];
			// definitionRelativeReferences.forEach((drr) => {
			// 	let actionSchema = schema.definitions[drr];
			// 	this.availableActionSchemas[drr] = actionSchema;

			// 	let actionVerb = this._getVerbFromActionSchema(actionSchema);
			// 	this.availableActions.push(actionVerb);
			// 	this.availableActionSchemasByVerb[actionVerb] = actionSchema;
			// 	// Do this only if current action has subactions definitions
			// 	// this.availableSubActionByVerb[actionVerb] = [];
			// 	// this.availableSubActionSchemasByVerb[actionVerb] = [];
			// });
		});
	}

	private _getSchema(): Promise<any> {
		return Promise.resolve(Schema);
	}

	private _getElementSchema(schema: any, object: any, property: string = null): any {
		let value = !property ? object : object[property];
		if (value['$ref']) {
			let definitionKey = value['$ref'].replace('#/definitions/', '');
			return schema.definitions[definitionKey];
		}

		return value;
	}

	private _getVerbFromActionSchema(actionDefinition: any): string {
		if (
			!actionDefinition.properties ||
			!actionDefinition.properties.verb ||
			!actionDefinition.properties.verb.enum ||
			!actionDefinition.properties.verb.enum.length
		) {
			throw new Error('Invalid Action schema');
		}

		return actionDefinition.properties.verb.enum[0];
	}

	private _getSubActionsSchemaFromParentActionSchema(schema: any, parentActionDefinition: any): any[] {
		if (!parentActionDefinition.properties) {
			throw new Error('Invalid Action schema');
		}

		if (!parentActionDefinition.properties.subactions) {
			return null;
		}

		if (
			parentActionDefinition.properties.subactions.type != 'array' ||
			!parentActionDefinition.properties.subactions.items ||
			!parentActionDefinition.properties.subactions.items.anyOf
		) {
			throw new Error('Invalid Action schema');
		}

		return parentActionDefinition.properties.subactions.items.anyOf.map((sa) => this._getElementSchema(schema, sa));
	}

	public getActionSchemaAsync(action: ISiteScriptAction): Promise<any> {
		return this.ensureInitialized().then(() => this.availableActionSchemas[action.verb]);
	}

	public getActionSchema(action: ISiteScriptAction): any {
		return this.availableActionSchemas[action.verb];
	}

	public getSubActionSchema(parentAction: ISiteScriptAction, subAction: ISiteScriptAction): any {
		return this.availableSubActionSchemasByVerb[parentAction.verb][subAction.verb];
	}

	public getAvailableActionsAsync(): Promise<string[]> {
		return this.ensureInitialized().then(() => this.availableActions);
	}

	public getAvailableActions(): string[] {
		return this.availableActions;
	}

	public getAvailableSubActions(parentAction: ISiteScriptAction): string[] {
		return this.availableSubActionByVerb[parentAction.verb];
  }

  public getAvailableSubActionsAsync(parentAction: ISiteScriptAction) : Promise<string[]> {
    return this.ensureInitialized().then(() => this.availableSubActionByVerb[parentAction.verb]);
  }
}

export const SiteScriptSchemaServiceKey = ServiceKey.create<ISiteScriptSchemaService>(
	'YPCODE:SiteScriptSchemaService',
	SiteScriptSchemaService
);
