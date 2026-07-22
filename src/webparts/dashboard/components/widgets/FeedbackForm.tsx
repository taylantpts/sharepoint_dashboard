import * as React from 'react';
import { TextField, PrimaryButton, MessageBar, MessageBarType, Spinner, SpinnerSize, useTheme, mergeStyleSets } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import { submitFeedback } from '../../services/SharePointService';
import { SUBMIT_UNAVAILABLE_MESSAGE } from '../../constants';

export interface IFeedbackFormProps {
    context: WebPartContext;
}

type SubmitState = 'idle' | 'sending' | 'success' | 'error';

const FeedbackForm: React.FunctionComponent<IFeedbackFormProps> = (props) => {
    const { context } = props;
    const theme = useTheme();
    const [message, setMessage] = React.useState('');
    const [state, setState] = React.useState<SubmitState>('idle');
    const [errorDetail, setErrorDetail] = React.useState<string | undefined>(undefined);

    const styles = mergeStyleSets({
        actions: {
            display: 'flex',
            justifyContent: 'flex-end',
            marginTop: 10
        },
        errorDetail: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginTop: 6,
            fontFamily: 'Consolas, monospace',
            wordBreak: 'break-word'
        }
    });

    const handleSubmit = async (): Promise<void> => {
        if (!message.trim()) {
            return;
        }
        setState('sending');
        setErrorDetail(undefined);

        try {
            const result = await submitFeedback(context, message.trim());
            if (result.success) {
                setState('success');
                setMessage('');
            } else {
                console.error('[FeedbackForm] Geri bildirim gönderilemedi:', result.errorMessage);
                setErrorDetail(result.errorMessage);
                setState('error');
            }
        } catch (error) {
            console.error('[FeedbackForm] Geri bildirim gönderilirken beklenmeyen hata:', error);
            setErrorDetail((error as Error).message);
            setState('error');
        }
    };

    return (
        <WidgetCard title="Geri Bildirim" subtitle="Görüş ve önerilerinizi bize iletin" iconName="Feedback" accentColor="#c77fae">
            {state === 'success' && (
                <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setState('idle')}>
                    Geri bildiriminiz için teşekkürler, iletildi.
                </MessageBar>
            )}
            {state === 'error' && (
                <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setState('idle')}>
                    {SUBMIT_UNAVAILABLE_MESSAGE}
                    {errorDetail && <div className={styles.errorDetail}>{errorDetail}</div>}
                </MessageBar>
            )}

            <TextField
                multiline
                rows={3}
                placeholder="Portal, süreçler veya çalışma ortamıyla ilgili görüşünüzü yazın..."
                value={message}
                onChange={(_, newValue) => setMessage(newValue ?? '')}
                disabled={state === 'sending'}
            />

            <div className={styles.actions}>
                {state === 'sending' ? (
                    <Spinner size={SpinnerSize.small} label="Gönderiliyor..." />
                ) : (
                    <PrimaryButton text="Gönder" onClick={() => { handleSubmit().catch(() => { /* handled in handleSubmit */ }); }} disabled={!message.trim()} />
                )}
            </div>
        </WidgetCard>
    );
};

export default FeedbackForm;
