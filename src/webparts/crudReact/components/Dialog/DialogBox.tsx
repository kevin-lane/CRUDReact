import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { useBoolean, useId } from '@fluentui/react-hooks';

export interface IDialogBoxProps {
    onConfirm: any;
    onCancel: any;
    hideDialog: boolean;
}

export const DialogBox: React.FunctionComponent<IDialogBoxProps> = (props: IDialogBoxProps) => {

  const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(false);
  const labelId: string = useId('dialogLabel');
  const subTextId: string = useId('subTextLabel');

    const dialogContentProps = {
        type: DialogType.normal,
        title: 'Delete item',
        closeButtonAriaLabel: 'Close',
        subText: 'Are you sure you want to delete list item ?',
      };

    const modalProps = React.useMemo(
        () => ({
          titleAriaId: labelId,
          subtitleAriaId: subTextId,
          isBlocking: false
        }),
        [labelId, subTextId],
      );

    return(
        <div>
            <Dialog
                hidden={props.hideDialog}
                onDismiss={toggleHideDialog}
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
            >
            <DialogFooter>
              <PrimaryButton onClick={props.onConfirm} text="Yes" />
              <DefaultButton onClick={props.onCancel} text="No" />
            </DialogFooter>
      </Dialog>
        </div>
    );
};