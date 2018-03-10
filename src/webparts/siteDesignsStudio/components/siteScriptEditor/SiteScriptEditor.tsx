import * as React from 'react';
import {
	Button,
	Dropdown,
	TextField,
	IconButton,
	Panel,
	PanelType,
	PrimaryButton,
	DefaultButton,
	CommandBar,
	MessageBar,
	MessageBarType,
	Spinner,
	SpinnerSize,
	DialogFooter,
	Pivot,
	PivotItem,
	CommandBarButton,
	IContextualMenuItem
} from 'office-ui-fabric-react';
import styles from '../SiteDesignsStudio.module.scss';
import { escape, assign } from '@microsoft/sp-lodash-subset';

import ScriptActionEditor from '../scriptActionEditor/ScriptActionEditor';
import { ISiteScriptContent, ISiteScriptAction, ISiteScript, SiteScriptEntitySchema } from '../../models/ISiteScript';
import Schema from '../../schema/schema';
import { IServiceConsumerComponentProps } from '../ISiteDesignsStudioProps';
import {
	ISiteScriptSchemaService,
	SiteScriptSchemaServiceKey
} from '../../services/siteScriptSchema/SiteScriptSchemaService';
import ScriptActionAdder from '../scriptActionAdder/ScriptActionAdder';
import { SiteDesignsServiceKey, ISiteDesignsService } from '../../services/siteDesigns/SiteDesignsService';
import GenericObjectEditor from '../genericObjectEditor/GenericObjectEditor';
import { MonacoEditor } from '../monacoEditor/MonacoEditor';
import ScriptActionCollectionEditor from '../scriptActionEditor/ScriptActionCollectionEditor';

const Ajv = require('ajv');
var ajv = new Ajv({ schemaId: 'auto' });

export enum EEditionMode {
	Design = 'design',
	Code = 'code',
	Split = 'split'
}

export interface ISiteScriptEditorState {
	script: ISiteScript;
	scriptContentJson: string;
	schema: any;
	isValidContent: boolean;
	isInvalidSchema: boolean;
	isNewScript: boolean;
	editMode: EEditionMode;
	isLoading: boolean;
	hasError: boolean;
	userMessage: string;
	expandedIndices: number[];
	allSubactionsExpanded: boolean;
	isEditingProperties: boolean;
}

export interface ISiteScriptEditorProps extends IServiceConsumerComponentProps {
	script: ISiteScript;
	onScriptUpdated?: (script: ISiteScript) => void;
}

export default class SiteScriptEditor extends React.Component<ISiteScriptEditorProps, ISiteScriptEditorState> {
	private siteScriptSchemaService: ISiteScriptSchemaService;
  private siteDesignsService: ISiteDesignsService;
  private stateHistory: ISiteScriptEditorState[];

	constructor(props: ISiteScriptEditorProps) {
		super(props);

		this.state = {
			script: null,
			scriptContentJson: '',
			schema: null,
			isNewScript: false,
			isValidContent: true,
			isInvalidSchema: false,
			editMode: EEditionMode.Split,
			isLoading: true,
			hasError: false,
			userMessage: '',
			expandedIndices: [],
			allSubactionsExpanded: false,
			isEditingProperties: false
		};
	}

	public componentWillMount() {
		this.props.serviceScope.whenFinished(() => {
			this.siteScriptSchemaService = this.props.serviceScope.consume(SiteScriptSchemaServiceKey);
			this.siteDesignsService = this.props.serviceScope.consume(SiteDesignsServiceKey);

			this._loadScript().then((loadedScript) => {
				const schema = this.siteScriptSchemaService.getSiteScriptSchema();
				// If the script content is not loaded => ERROR
				if (!loadedScript.Content && loadedScript.Id) {
					this.setState({
						script: null,
						scriptContentJson: 'INVALID SCHEMA',
						schema: schema,
						isLoading: false,
						isValidContent: false,
						hasError: true,
						userMessage: 'The specified script is invalid'
					});
					return;
				}

				this.setState({
					script: loadedScript,
					schema: schema,
					isNewScript: loadedScript.Id ? false : true,
					isLoading: false,
					isInvalidSchema: false,
					scriptContentJson: JSON.stringify(loadedScript.Content, null, 2)
				});
			});
		});
	}

	public componentDidUpdate() {
		this._autoClearInfoMessages();
	}

	private _loadScript(): Promise<ISiteScript> {
		let { script } = this.props;

		// If existing script (The Id is known)
		if (script.Id) {
			// Load that script
			return this.siteDesignsService.getSiteScript(script.Id);
		} else {
			// If the argument is a new script
			// Initialize the content
			return this._initializeScriptContent(script);
		}
	}

	public render(): React.ReactElement<ISiteScriptEditorProps> {
		let { isLoading, isEditingProperties, script, isNewScript, hasError, userMessage } = this.state;

		if (isLoading) {
			return (
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm6 ms-smOffset3">
						<Spinner size={SpinnerSize.large} label="Loading..." />
					</div>
				</div>
			);
		}

		return (
			<div>
				{userMessage && (
					<MessageBar
						className={userMessage ? 'ms-fadeIn400' : 'ms-fadeOut400'}
						messageBarType={hasError ? MessageBarType.error : MessageBarType.success}
					>
						{userMessage}
					</MessageBar>
				)}

				{(isNewScript || isEditingProperties) && this._renderSiteScriptPropertiesEditor()}
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12">
						<CommandBar items={this._getCommands()} farItems={this._getFarCommands()} />
					</div>
				</div>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12">
						<div className={styles.designWorkspace}>{this._renderEditor()}</div>
					</div>
				</div>
			</div>
		);
	}

	private _renderSiteScriptPropertiesEditor() {
		let { script } = this.state;
		let editingSiteScript = assign({}, script);

		const onObjectChanged = (o) => {
			assign(editingSiteScript, o);
		};

		// If the Id is not set, do not render it
		let ignoredProperties = editingSiteScript.Id ? [] : [ 'Id' ];

		return (
			<Panel isOpen={true} type={PanelType.smallFixedFar} onDismiss={() => this._cancelScriptPropertiesEdition()}>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12">
						<GenericObjectEditor
							ignoredProperties={ignoredProperties}
							readOnlyProperties={[ 'Id' ]}
							object={editingSiteScript}
							onObjectChanged={onObjectChanged.bind(this)}
							schema={SiteScriptEntitySchema}
						/>
					</div>
				</div>
				<DialogFooter>
					<PrimaryButton text="Ok" onClick={() => this._applyPropertiesEdition(editingSiteScript)} />
					<DefaultButton text="Cancel" onClick={() => this._cancelScriptPropertiesEdition()} />
				</DialogFooter>
			</Panel>
		);
	}

	private _renderEditor() {
		let {
			script,
			schema,
			scriptContentJson,
			isInvalidSchema,
			isValidContent,
			allSubactionsExpanded,
			expandedIndices
		} = this.state;

		const codeEditor = (
			<div className="ms-Grid-row">
				<div className="ms-Grid-col ms-sm12">
					<div className={styles.codeEditorContainer}>
						<MonacoEditor
							schema={schema}
							value={scriptContentJson}
							onValueChange={(content, errors) => this._onCodeUpdated(content, errors)}
							readOnly={false}
						/>
					</div>
				</div>
			</div>
		);

		const designerEditor = isValidContent && (
			<div className="ms-Grid-row">
				<div className="ms-Grid-col ms-sm12">
					<div className="ms-Grid-row">
						<ScriptActionCollectionEditor
							serviceScope={this.props.serviceScope}
							actions={script.Content.actions as ISiteScriptAction[]}
							onActionRemoved={(actionIndex) => this._removeScriptAction(actionIndex)}
							onActionMoved={(oldIndex, newIndex) => this._moveAction(oldIndex, newIndex)}
							onActionChanged={(actionIndex, action) => this._onActionUpdated(actionIndex, action)}
							expandedIndices={expandedIndices}
							onExpandChanged={(expanded) => this._onExpandChanged(expanded)}
							getActionSchema={(action) => this.siteScriptSchemaService.getActionSchema(action)}
						/>
					</div>
					<div className="ms-Grid-row">
						<div>
							<ScriptActionAdder
								serviceScope={this.props.serviceScope}
								onActionAdded={(a) => this._addScriptAction(a)}
							/>
						</div>
					</div>
				</div>
			</div>
		);

		switch (this.state.editMode) {
			case EEditionMode.Code:
				return codeEditor;
			case EEditionMode.Design:
				return designerEditor;
			case EEditionMode.Split:
			default:
				return (
					<div className={styles.splitWorkspace}>
						<div className={styles.splitPane}>{designerEditor}</div>
						<div className={styles.splitPane}>{codeEditor}</div>
					</div>
				);
		}
	}

	private _initializeScriptContent(script: ISiteScript): Promise<ISiteScript> {
		script.Content = this.siteScriptSchemaService.getNewSiteScript();
		return Promise.resolve(script);
	}

	private _getCommands() {
		let { script, editMode, isValidContent, expandedIndices } = this.state;
    let actionsCount = script.Content.actions.length;
    const undoBtn: IContextualMenuItem = {
			key: 'undoBtn',
			text: 'Undo',
      title: 'Undo',
      disabled: !this._canUndo(),
			iconProps: { iconName: 'Undo' },
			onClick: () => this._undo()
		};
		const saveBtn: IContextualMenuItem = {
			key: 'saveBtn',
			text: 'Save',
			title: 'Save',
			iconProps: { iconName: 'Save' },
			onClick: () => this._saveSiteScript()
		};
		const editBtn: IContextualMenuItem = {
			key: 'btnEdit',
			text: 'Edit Properties',
			title: 'Edit Properties',
			iconProps: { iconName: 'Edit' },
			onClick: () => this._editProperties()
		};
		const expandAllBtn: IContextualMenuItem = {
			key: 'expandAllBtn',
			text: 'Expand All',
			title: 'Expand All',
			disabled: expandedIndices.length == actionsCount,
			iconProps: { iconName: 'ExploreContent' },
			onClick: () => this._setAllExpanded(true)
		};
		const collapseAllBtn: IContextualMenuItem = {
			key: 'btnCollapseAll',
			text: 'Collapse All',
			title: 'Collapse All',
			disabled: expandedIndices.length == 0,
			iconProps: { iconName: 'CollapseContent' },
			onClick: () => this._setAllExpanded(false)
		};

		let commands = [undoBtn];

		if (isValidContent) {
			commands = commands.concat(saveBtn, editBtn);

			if (editMode != EEditionMode.Code) {
				commands = commands.concat(expandAllBtn, collapseAllBtn);
			}
		}

		return commands;
	}

	private _getFarCommands() {
		let { script, editMode, isValidContent, expandedIndices } = this.state;
		let actionsCount = script.Content.actions.length;

		const designModeButton: IContextualMenuItem = {
			key: 'designModeBtn',
			text: 'Design',
			title: 'Design',
			iconProps: { iconName: 'Design' },
			onClick: () => this._setEditionMode(EEditionMode.Design)
		};

		const codeModeButton: IContextualMenuItem = {
			key: 'codeModeBtn',
			text: 'Code',
			title: 'Code',
			iconProps: { iconName: 'Code' },
			onClick: () => this._setEditionMode(EEditionMode.Code)
		};

		const splitModeButton: IContextualMenuItem = {
			key: 'splitModeBtn',
			text: 'Split',
			title: 'Split',
			iconProps: { iconName: 'Split' },
			onClick: () => this._setEditionMode(EEditionMode.Split)
		};

		return [ designModeButton, codeModeButton, splitModeButton ];
	}

	private _setAllExpanded(isExpanded: boolean) {
		let { expandedIndices, script } = this.state;
		let actions = script.Content.actions;
		this.setState({
			expandedIndices: isExpanded ? actions.map((item, index) => index) : []
		});
	}

	private _setEditionMode(mode: EEditionMode) {
		this.setState({ editMode: mode });
	}

	private _editProperties() {
		this.setState({
			isEditingProperties: true
		});
	}

	private _onExpandChanged(expandedIndices: number[]) {
		this.setState({
			expandedIndices: expandedIndices
		});
	}
	private _moveAction(oldIndex: number, newIndex: number) {
		let { script, expandedIndices } = this.state;
		if (newIndex < 0 || newIndex > script.Content.actions.length - 1) {
			return;
		}

		let newActions = [].concat(script.Content.actions);
		let actionToMove = newActions[oldIndex];
		// If the moved item is expanded
		if (expandedIndices.indexOf(oldIndex) > -1) {
			expandedIndices = expandedIndices.filter((a) => a != oldIndex).concat(newIndex);
		}
		newActions.splice(oldIndex, 1);
		newActions.splice(newIndex, 0, actionToMove);
		let newContent = assign({}, script.Content);
		newContent.actions = newActions;
		let newScript = assign({}, script);
		newScript.Content = newContent;

    this._saveToStateHistory();
		this.setState({
			script: newScript,
			scriptContentJson: JSON.stringify(newScript.Content, null, 2),
			expandedIndices: expandedIndices
		});
	}

	private _addScriptAction(verb: string) {
		let { script } = this.state;
		let newAction: ISiteScriptAction = {
			verb: verb
		};

		let newActionsArray = [].concat(script.Content.actions, newAction);
		let newScriptContent = assign({}, script.Content);
		newScriptContent.actions = newActionsArray;
		let newScript = assign({}, script);
    newScript.Content = newScriptContent;
    this._saveToStateHistory();
		this.setState({
			script: newScript,
			scriptContentJson: JSON.stringify(newScript.Content, null, 2),
			expandedIndices: [ newActionsArray.length - 1 ]
		});
	}

	private _removeScriptAction(actionIndex: number) {
		let { script } = this.state;
		let newActionsArray = script.Content.actions.filter((item, index) => index != actionIndex);
		let newScriptContent = assign({}, script.Content);
		newScriptContent.actions = newActionsArray;
		let newScript = assign({}, script);
    newScript.Content = newScriptContent;
    this._saveToStateHistory();
		this.setState({
			script: newScript,
			scriptContentJson: JSON.stringify(newScript.Content, null, 2)
		});
	}

	private _onActionUpdated(actionKey: number, action: ISiteScriptAction) {
		let { script } = this.state;
		let newScript: ISiteScript = assign({}, script);
		let newScriptContent = assign({}, script.Content);

		newScriptContent.actions = [].concat(newScriptContent.actions);

		// Replace the appropriate action
		newScriptContent.actions[actionKey] = action;

    newScript.Content = newScriptContent;

    this._saveToStateHistory();
		this.setState({
			script: newScript,
			scriptContentJson: JSON.stringify(newScript.Content, null, 2)
		});
	}

	private _onCodeUpdated(code: string, errors: any) {
		console.log('Erros: ', errors);

		let { script, schema } = this.state;

		// Validate the schema
		let parsedCode = JSON.parse(code);
		console.log('Object to validate: ', parsedCode);
		let valid = ajv.validate(schema, parsedCode);
		if (!valid) {
			console.log('Schema is not valid');
			console.log('Schema errors: ', ajv.errors);

			this.setState({
				scriptContentJson: code,
				isLoading: false,
				isValidContent: false,
				hasError: true,
				userMessage: 'Oops... The Site Script is invalid...'
			});
		} else {
			let newScript: ISiteScript = assign({}, script);

			let newScriptContent = parsedCode;

      newScript.Content = newScriptContent;

			this.setState({
				script: newScript,
				scriptContentJson: JSON.stringify(newScript.Content, null, 2),
				isValidContent: true,
				hasError: false,
				userMessage: null
			});
		}
	}

	private _validateForSave(): string {
		let { schema, script } = this.state;

		if (!script.Title) {
			return 'The Site Script has no title';
		}

		if (!script.Content) {
			return 'The Site Script has no content';
		}

		// Check content schema validity
		let valid = ajv.validate(schema, script.Content);
		if (!valid) {
			return 'The Site Script is not valid against the Schena';
		}

		return null;
	}
	private _saveSiteScript() {
		let { script } = this.state;
		let { onScriptUpdated } = this.props;
		let invalidMessage = this._validateForSave();
		if (invalidMessage) {
			this.setState({
				hasError: true,
				userMessage: invalidMessage
			});
			return;
		}

		this.setState({ isLoading: true, isEditingProperties: false });
		this.siteDesignsService
			.saveSiteScript(script)
			.then((_) => {
				this.setState({
					isEditingProperties: false,
					isNewScript: false,
					isLoading: false,
					hasError: false,
					userMessage: 'The site script have been properly saved'
        });
        this._clearStateHistory();
			})
			.catch((error) => {
				this.setState({
					isEditingProperties: false,
					hasError: true,
					isNewScript: false,
					isLoading: false,
					userMessage: 'The site script cannot be properly saved'
				});
			});
	}

	private _autoClearInfoMessages() {
		let { userMessage, hasError } = this.state;
		if (hasError) {
			return;
		}

		if (userMessage) {
			setTimeout(() => {
				this.setState({
					userMessage: null
				});
			}, 3000);
		}
	}

	private _applyPropertiesEdition(siteScript: ISiteScript) {
		this.setState({
			script: siteScript,
			scriptContentJson: JSON.stringify(siteScript.Content, null, 2),
			isEditingProperties: false,
			isNewScript: false
		});

		if (this.props.onScriptUpdated) {
			this.props.onScriptUpdated(siteScript);
		}
	}
	private _cancelScriptPropertiesEdition() {
		this.setState({
			isEditingProperties: false,
			isNewScript: false
		});
  }

  private _saveToStateHistory() {
    if (!this.stateHistory) {
      this.stateHistory = [];
    }

    if (this.stateHistory.length > 10) {
      this.stateHistory.splice(9, 1);
    }

    this.stateHistory.splice(0,0, this.state);
  }

private _clearStateHistory() {
  this.stateHistory = null;
}

  private _undo() {
    if (!this.stateHistory || this.stateHistory.length == 0) {
      return;
    }

    let previousState = this.stateHistory[0];
    this.stateHistory.splice(0, 1);
    this.setState(previousState);
  }

  private _canUndo() : boolean {
    return this.stateHistory && this.stateHistory.length > 0;
  }
}
