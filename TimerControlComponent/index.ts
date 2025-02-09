import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class TimerControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private timerDisplay: HTMLDivElement;
    private startStopButton: HTMLButtonElement;
    private timerInterval: NodeJS.Timeout | undefined;
    private startTime: Date | null = null;
    private durationInSeconds: number = 0;
    private context: ComponentFramework.Context<IInputs>;
    private recordId: string | null = null;

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this.context = context;
        this.container = container;
        this.recordId = context.parameters.EntityId.raw || null; // Get the record ID

        this.container.classList.add("timer-container");

        // Create start/stop button
        this.startStopButton = document.createElement("button");
        this.startStopButton.id = "startStopButton";
        this.startStopButton.textContent = "Start";
        this.startStopButton.classList.add("start-stop-button");
        this.startStopButton.addEventListener("click", this.toggleTimer.bind(this));
        this.container.appendChild(this.startStopButton);

        // Create timer display
        this.timerDisplay = document.createElement("div");
        this.timerDisplay.id = "timer";
        this.timerDisplay.textContent = "00:00:00";
        this.timerDisplay.classList.add("timer-display");
        this.container.appendChild(this.timerDisplay);

        this.loadTimerState(); // Load timer state on initialization
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        const newRecordId = context.parameters.EntityId.raw || null;
        if (newRecordId !== this.recordId) {
            this.recordId = newRecordId;
            this.loadTimerState(); // Reload timer state when switching records
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        if (this.timerInterval) {
            clearInterval(this.timerInterval);
        }
        this.startStopButton.removeEventListener("click", this.toggleTimer.bind(this));
    }

    private toggleTimer(): void {
        if (this.startStopButton.textContent === "Start") {
            this.startTimer();
        } else {
            this.stopTimer();
        }
    }

    private startTimer(): void {
        if (!this.recordId) return;

        this.startStopButton.textContent = "Stop";
        this.startTime = new Date();
        this.saveTimerState(true); // Save timer state when starting

        // Update timer every second
        this.timerInterval = setInterval(() => {
            if (this.startTime) {
                const elapsed = Math.floor((new Date().getTime() - this.startTime.getTime()) / 1000);
                this.durationInSeconds = elapsed;
                this.timerDisplay.textContent = this.formatTime(elapsed);
            }
        }, 1000);
    }

    private stopTimer(): void {
        if (!this.recordId) return;

        if (this.timerInterval) {
            clearInterval(this.timerInterval);
        }
        this.startStopButton.textContent = "Start";

        const entityName = this.context.parameters.EntityName.raw; // Get entity name
        const fieldName = "sonade_timespent"; // Field to update in Dataverse

        if (!entityName) {
            console.error("No entity name found.");
            return;
        }

        const updateData: any = {};
        updateData[fieldName] = this.durationInSeconds; // Set time spent value

        // Update the record field with the elapsed time
        this.context.webAPI.updateRecord(entityName, this.recordId, updateData)
            .then(() => {
                console.log(`Field '${fieldName}' updated successfully.`);
            })
            .catch((error: any) => {
                console.error(`Error updating field '${fieldName}':`, error);
            });

        this.resetTimer();
    }

    private resetTimer(): void {
        if (!this.recordId) return;

        this.startTime = null;
        this.durationInSeconds = 0;
        this.timerDisplay.textContent = "00:00:00";
        this.saveTimerState(false); // Reset timer state when stopped
    }

    private loadTimerState(): void {
        if (!this.recordId) return;

        const savedState = localStorage.getItem(`timer_${this.recordId}`);
        if (savedState) {
            const { startTime, running } = JSON.parse(savedState);
            if (running && startTime) {
                this.startTime = new Date(startTime);
                this.startStopButton.textContent = "Stop";

                // Resume the timer from stored state
                this.timerInterval = setInterval(() => {
                    if (this.startTime) {
                        const elapsed = Math.floor((new Date().getTime() - this.startTime.getTime()) / 1000);
                        this.durationInSeconds = elapsed;
                        this.timerDisplay.textContent = this.formatTime(elapsed);
                    }
                }, 1000);
            }
        }
    }

    private saveTimerState(isRunning: boolean): void {
        if (!this.recordId) return;

        const state = JSON.stringify({
            startTime: this.startTime ? this.startTime.toISOString() : null,
            running: isRunning,
        });
        localStorage.setItem(`timer_${this.recordId}`, state);
    }

    private formatTime(seconds: number): string {
        const hours = Math.floor(seconds / 3600);
        const minutes = Math.floor((seconds % 3600) / 60);
        const secs = seconds % 60;
        return `${this.pad(hours)}:${this.pad(minutes)}:${this.pad(secs)}`;
    }

    private pad(num: number): string {
        return num < 10 ? "0" + num : num.toString();
    }
}
