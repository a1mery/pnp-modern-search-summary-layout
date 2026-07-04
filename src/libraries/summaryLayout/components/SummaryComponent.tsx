import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { ServiceScope } from '@microsoft/sp-core-library';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';

export const DEFAULT_COPILOT_PROMPT = 'Provide a concise summary of this document. Focus on the key points and main takeaways. Avoid including minor details or irrelevant information. Use bullet points if it helps clarity.';

export class DocumentSummaryWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback(): Promise<void> {
        const props = this.resolveAttributes();
        const serviceScope: ServiceScope = this._serviceScope;
        const customComponent = <DocumentSummaryButton serviceScope={serviceScope} {...props} />;
        ReactDOM.render(customComponent, this);
    }

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}

export interface IDocumentSummaryButtonProps {
    serviceScope: ServiceScope;
    documentUrl?: string;
    title?: string;
    promptText?: string;
}

export interface IDocumentSummaryButtonState {
    showDialog: boolean;
    isLoading: boolean;
    summary: string | null;
    error: string | null;
}

export class DocumentSummaryButton extends React.Component<IDocumentSummaryButtonProps, IDocumentSummaryButtonState> {

    public constructor(props: IDocumentSummaryButtonProps) {
        super(props);
        this.state = {
            showDialog: false,
            isLoading: false,
            summary: null,
            error: null
        };
    }

    /**
     * Converts Copilot markdown response to safe HTML.
     * Handles: bold, italic, links, citations [n](url), line breaks, and paragraphs.
     */
    private markdownToHtml(md: string): string {
        let html = md;

        // Escape HTML entities to prevent XSS
        html = html
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');

        // Bold: **text**
        html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');

        // Italic: *text* (but not inside bold markers)
        html = html.replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, '<em>$1</em>');

        // Citation references: [n](url) — render as superscript link
        html = html.replace(/\[(\d+)\]\((https?:\/\/[^\s)]+)\)/g,
            '<a href="$2" target="_blank" rel="noopener noreferrer" style="color:#0078d4;text-decoration:none;font-size:11px;vertical-align:super">[$1]</a>');

        // Inline links: [text](url)
        html = html.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g,
            '<a href="$2" target="_blank" rel="noopener noreferrer" style="color:#0078d4;text-decoration:none">$1</a>');

        // Paragraphs: double newlines
        html = html.replace(/\n\n+/g, '</p><p>');

        // Single line breaks
        html = html.replace(/\n/g, '<br/>');

        // Wrap in paragraph tags
        html = '<p>' + html + '</p>';

        return html;
    }

    private handleSummarizeClick = (): void => {
        this.setState({
            showDialog: true,
            isLoading: true,
            summary: null,
            error: null
        });
        this.fetchSummary();
    }

    private fetchSummary = async (): Promise<void> => {
        try {
            const { serviceScope, documentUrl } = this.props;

            if (!documentUrl) {
                this.setState({
                    error: 'No document URL available.',
                    isLoading: false
                });
                return;
            }

            const msGraphClientFactory = serviceScope.consume<MSGraphClientFactory>(MSGraphClientFactory.serviceKey);
            const graphClient = await msGraphClientFactory.getClient('3');

            // Step 1: Create a Copilot conversation
            const conversationResponse = await graphClient
                .api('https://graph.microsoft.com/beta/copilot/conversations')
                .post({});

            const conversationId: string = conversationResponse.id;

            // Step 2: Send a message asking to summarize the document
            const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
            const promptText = (this.props.promptText && this.props.promptText.trim().length > 0)
                ? this.props.promptText
                : DEFAULT_COPILOT_PROMPT;
            const chatResponse = await graphClient
                .api(`https://graph.microsoft.com/beta/copilot/conversations/${conversationId}/chat`)
                .post({
                    message: {
                        text: promptText,
                    },
                    locationHint: {
                        timeZone: timeZone
                    },
                    contextualResources: {
                        files: [
                            {
                                uri: documentUrl
                            }
                        ],
                        webContext: {
                            isWebEnabled: false
                        }
                    }
                });

            // Extract the response text from the last message
            const messages: Array<{ text: string }> = chatResponse.messages || [];
            const responseMessage = messages.length > 1 ? messages[messages.length - 1] : null;
            const summaryText = responseMessage ? responseMessage.text : 'No summary was generated.';

            this.setState({
                summary: summaryText,
                isLoading: false
            });

        } catch (err: unknown) {
            console.error('Error fetching document summary:', err);

            let errorMessage = 'An unexpected error occurred.';
            if (err instanceof Error) {
                errorMessage = err.message;
            } else if (typeof err === 'object' && err !== null) {
                // Graph SDK errors often have a body or statusCode property
                const graphErr = err as Record<string, unknown>;
                if (graphErr.statusCode && graphErr.message) {
                    errorMessage = `HTTP ${graphErr.statusCode}: ${graphErr.message}`;
                } else if (graphErr.body) {
                    errorMessage = typeof graphErr.body === 'string' ? graphErr.body : JSON.stringify(graphErr.body);
                } else {
                    errorMessage = JSON.stringify(err);
                }
            }

            this.setState({
                error: errorMessage,
                isLoading: false
            });
        }
    }

    private closeDialog = (): void => {
        if (!this.state.isLoading) {
            this.setState({
                showDialog: false,
                summary: null,
                error: null
            });
        }
    }

    public render(): React.ReactElement<IDocumentSummaryButtonProps> {
        const { showDialog, isLoading, summary, error } = this.state;
        const { title } = this.props;

        return (
            <div>
                <PrimaryButton
                    onClick={this.handleSummarizeClick}
                    styles={{
                        root: {
                            minWidth: 'auto',
                            padding: '4px 12px',
                            height: '32px'
                        }
                    }}
                >
                    <Icon iconName="Robot" styles={{ root: { marginRight: 6 } }} />
                    Summarize
                </PrimaryButton>

                <Dialog
                    hidden={!showDialog}
                    onDismiss={this.closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: `Document Summary`,
                        subText: title ? `Summary for: ${title}` : undefined
                    }}
                    minWidth={550}
                    modalProps={{
                        isBlocking: isLoading
                    }}
                >
                    {isLoading ? (
                        <div style={{ padding: '30px 0', textAlign: 'center' }}>
                            <Spinner label="Generating summary with Copilot..." size={SpinnerSize.large} />
                        </div>
                    ) : error ? (
                        <MessageBar messageBarType={MessageBarType.error}>
                            {error}
                        </MessageBar>
                    ) : summary ? (
                        <div
                            style={{
                                padding: '10px 0',
                                lineHeight: '1.6',
                                fontSize: '14px',
                                color: '#323130',
                                maxHeight: '400px',
                                overflowY: 'auto'
                            }}
                            dangerouslySetInnerHTML={{ __html: this.markdownToHtml(summary) }}
                        />
                    ) : null}
                    <DialogFooter>
                        <DefaultButton onClick={this.closeDialog} text="Close" disabled={isLoading} />
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }
}
