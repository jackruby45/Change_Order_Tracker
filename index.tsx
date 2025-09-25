/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/
import { render } from "preact";
// FIX: Import JSX type to resolve 'preact.JSX' namespace errors.
import type { JSX } from "preact";
import { useState, useMemo, useEffect, useRef } from "preact/hooks";
import { html } from "htm/preact";
import jsPDF from "jspdf";
import autoTable from 'jspdf-autotable';

// --- TYPES ---
type ApprovalStatus = 'Pending' | 'Approved' | 'Rejected';
type ChangeOrderStatus = 'Pending Approval' | 'Approved' | 'Rejected' | 'In Progress' | 'Completed';

interface Approval {
    name: string;
    status: ApprovalStatus;
    approvalDate: string | null;
}

interface ChangeOrder {
    id: number;
    title: string;
    description: string;
    reason: string;
    status: ChangeOrderStatus;
    dateRequested: string;
    costImpactEquipment: number;
    costImpactInstallation: number;
    costImpactOther: number;
    otherCostsExplanation: string;
    scheduleImpactDays: number;
    approvals: Approval[];
}

type ApproverConfig = { [role: string]: string };

interface AppSettings {
    projectName: string;
    projectLocation: string;
    projectManager: string;
    approverConfig: ApproverConfig;
}

const APPROVER_ROLES = [
    'Manager of Gas Engineering',
    'Director of Gas Engineering',
    'Director of Gas Operations',
    'Sr. Vice President of Gas Operations'
];

// --- UTILITY FUNCTIONS ---
const formatDate = (dateString: string | null) => dateString ? new Date(dateString).toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric', timeZone: 'UTC' }) : 'N/A';
const formatCurrency = (amount: number) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(amount);
const getStatusColor = (status: ChangeOrderStatus | ApprovalStatus) => {
    switch(status) {
        case 'Pending':
        case 'Pending Approval': return 'var(--status-pending)';
        case 'Approved': return 'var(--status-approved)';
        case 'Rejected': return 'var(--status-rejected)';
        case 'In Progress': return 'var(--status-inprogress)';
        case 'Completed': return 'var(--status-completed)';
        default: return 'var(--text-secondary)';
    }
};

// --- PDF GENERATION ---
const generatePdfReport = (changeOrders: ChangeOrder[], projectName: string, projectLocation: string, projectManager: string) => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const margin = 15;

    // --- TITLE PAGE ---
    doc.setFontSize(32);
    doc.text('UNITIL', pageWidth / 2, 60, { align: 'center' });
    doc.setFontSize(26);
    doc.text('Change Order Report', pageWidth / 2, 80, { align: 'center' });

    const projectDetails = [
        ['Project Name:', projectName || 'Not Specified'],
        ['Project Location:', projectLocation || 'Not Specified'],
        ['Project Manager:', projectManager || 'Not Specified'],
        ['Report Date:', new Date().toLocaleDateString()]
    ];

    autoTable(doc, {
        body: projectDetails,
        startY: 100,
        theme: 'plain',
        styles: { fontSize: 12, cellPadding: 3 },
        columnStyles: {
            0: { fontStyle: 'bold' },
        },
        // Center the table on the page
        tableWidth: 'auto',
        margin: { left: margin },
    });


    // --- SUMMARY SECTION ---
    if (changeOrders.length > 0) {
        doc.addPage();
        let yPos = margin;
        doc.setFontSize(18);
        doc.text('Report Summary', margin, yPos);
        yPos += 10;
        doc.setLineWidth(0.5);
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 10;

        // --- Calculations ---
        const getTotalCost = (orders: ChangeOrder[]) => orders.reduce((sum, co) => sum + co.costImpactEquipment + co.costImpactInstallation + co.costImpactOther, 0);

        const approvedOrders = changeOrders.filter(co => ['Approved', 'In Progress', 'Completed'].includes(co.status));
        const pendingOrders = changeOrders.filter(co => co.status === 'Pending Approval');
        const rejectedOrders = changeOrders.filter(co => co.status === 'Rejected');
        
        const totalCostApproved = getTotalCost(approvedOrders);
        const totalCostPending = getTotalCost(pendingOrders);
        const totalCostRejected = getTotalCost(rejectedOrders);
        const totalCostAll = changeOrders.length > 0 ? getTotalCost(changeOrders) : 0;

        const totalDaysApproved = approvedOrders.reduce((sum, co) => sum + co.scheduleImpactDays, 0);

        // --- Summary Tables ---
        const summaryBody = [
            ['Total Change Orders', changeOrders.length.toString()],
            ['Total Schedule Impact (Approved)', `${totalDaysApproved} days`],
        ];

        const costSummaryBody = [
            ['Total Cost (All Orders)', formatCurrency(totalCostAll)],
            [' - Approved / In Progress', formatCurrency(totalCostApproved)],
            [' - Pending Approval', formatCurrency(totalCostPending)],
            [' - Rejected', formatCurrency(totalCostRejected)],
        ];

        autoTable(doc, {
            startY: yPos,
            body: summaryBody,
            theme: 'plain',
            styles: { fontSize: 12, cellPadding: 1.5 },
            columnStyles: {
                0: { fontStyle: 'bold' },
                1: { halign: 'right' }
            }
        });
        
        yPos = (doc as any).lastAutoTable.finalY + 8;
        
        doc.setFontSize(14);
        doc.text('Cost Summary', margin, yPos);
        yPos += 6;
        
        autoTable(doc, {
            startY: yPos,
            body: costSummaryBody,
            theme: 'plain',
            styles: { fontSize: 12, cellPadding: 1.5 },
            columnStyles: {
                0: { fontStyle: 'bold' },
                1: { halign: 'right' }
            }
        });
    }


    // --- DETAILED SECTION ---
    changeOrders.forEach(order => {
        doc.addPage();
        let currentY = margin;

        // --- Table 1 (Main Details) ---
        const totalCostOrder = order.costImpactEquipment + order.costImpactInstallation + order.costImpactOther;
        const mainDetailsData = [
            ['CO #', order.id],
            ['Title', order.title],
            ['Status', order.status],
            ['Date Requested', formatDate(order.dateRequested)],
            ['Description', order.description],
            ['Reason for Change', order.reason],
            ['Total Cost Impact', formatCurrency(totalCostOrder)],
            ['Schedule Impact', `${order.scheduleImpactDays} days`],
        ];

        if (order.otherCostsExplanation?.trim()) {
            mainDetailsData.push(['Other Costs Explanation', order.otherCostsExplanation]);
        }

        autoTable(doc, {
            startY: currentY,
            head: [['Field', 'Value']],
            body: mainDetailsData,
            theme: 'striped',
            headStyles: { fillColor: [42, 42, 42] },
            columnStyles: {
                0: { fontStyle: 'bold', cellWidth: 50 },
            },
            didDrawCell: (data) => {
                if (data.section === 'body' && data.column.index === 1) {
                    // FIX: The raw cell value can be a number (e.g., CO #), which does not have a `length` property.
                    // Convert to string before checking length to handle all data types safely.
                   if (String(data.cell.raw).length > 50) { // Arbitrary length for wrapping
                     // You could add custom logic here if needed, but autoTable handles wrapping well
                   }
                }
            }
        });
        
        currentY = (doc as any).lastAutoTable.finalY + 10;
        
        // --- Table 2 (Approval Details) ---
        doc.setFontSize(14);
        doc.text('Approval Details', margin, currentY);
        currentY += 8;

        autoTable(doc, {
            startY: currentY,
            head: [['Approver Name', 'Status', 'Approval Date']],
            body: order.approvals.length > 0
                ? order.approvals.map(app => [app.name, app.status, formatDate(app.approvalDate)])
                : [['No approvers assigned.']],
            theme: 'grid',
            headStyles: { fillColor: [42, 42, 42] },
        });
    });

    // --- PAGE NUMBERING ---
    // FIX: The `getNumberOfPages` method is part of the public jsPDF API and should be called on the `doc` object directly, not on `doc.internal`. 
    // This resolves the TypeScript error because the method is correctly defined on the `jsPDF` instance type.
    const pageCount = doc.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(10);
        doc.setTextColor(150);
        doc.text(
            `Page ${i} of ${pageCount}`,
            pageWidth / 2,
            pageHeight - 10,
            { align: 'center' }
        );
    }

    // --- SAVE FILE ---
    doc.save('Change_Order_Report.pdf');
};


// --- COMPONENTS ---

const ConfirmModal = ({ isOpen, onClose, onConfirm, title, children, confirmText = 'Confirm', confirmClass = 'btn-primary' }) => {
    if (!isOpen) return null;

    return html`
        <div class="modal-overlay" onClick=${onClose}>
            <div class="modal-content" onClick=${e => e.stopPropagation()}>
                <div class="modal-header">
                    <h2>${title}</h2>
                    <button class="btn-close" onClick=${onClose}>×</button>
                </div>
                <div class="modal-body">
                    ${children}
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" onClick=${onClose}>Cancel</button>
                    <button type="button" class="btn ${confirmClass}" onClick=${onConfirm}>${confirmText}</button>
                </div>
            </div>
        </div>
    `;
};

const PasswordModal = ({ isOpen, onClose, onConfirm, title }) => {
    if (!isOpen) return null;
    const [password, setPassword] = useState('');
    const inputRef = useRef<HTMLInputElement>(null);

    const handleSubmit = () => {
        onConfirm(password);
        setPassword('');
    };
    
    const handleKeyDown = (e) => {
        if (e.key === 'Enter') {
            handleSubmit();
        }
    };

    useEffect(() => {
        if (isOpen) {
            setTimeout(() => inputRef.current?.focus(), 0);
        }
    }, [isOpen]);

    return html`
        <div class="modal-overlay" onClick=${onClose}>
            <div class="modal-content" onClick=${e => e.stopPropagation()}>
                <div class="modal-header">
                    <h2>${title}</h2>
                    <button class="btn-close" onClick=${onClose}>×</button>
                </div>
                <div class="modal-body">
                    <div class="form-group">
                        <label for="admin-password">Password</label>
                        <input 
                            ref=${inputRef}
                            id="admin-password" 
                            type="password" 
                            class="form-input" 
                            value=${password} 
                            onInput=${e => setPassword(e.target.value)}
                            onKeyDown=${handleKeyDown}
                        />
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" onClick=${onClose}>Cancel</button>
                    <button type="button" class="btn btn-primary" onClick=${handleSubmit}>Submit</button>
                </div>
            </div>
        </div>
    `;
};

const App = () => {
    const [changeOrders, setChangeOrders] = useState<ChangeOrder[]>([]);
    const [currentOrder, setCurrentOrder] = useState<ChangeOrder | null>(null);
    const [view, setView] = useState<'table' | 'form'>('table');
    const [isSettingsOpen, setIsSettingsOpen] = useState(false);
    const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = useState(false);
    const [isRenumberConfirmOpen, setIsRenumberConfirmOpen] = useState(false);
    const [isPasswordModalOpen, setIsPasswordModalOpen] = useState(false);
    const [isAdminMode, setIsAdminMode] = useState(false);
    const [selectedOrderIds, setSelectedOrderIds] = useState(new Set<number>());
    const [settings, setSettings] = useState<AppSettings>({
        projectName: '',
        projectLocation: '',
        projectManager: '',
        approverConfig: {}
    });
    const fileInputRef = useRef<HTMLInputElement>(null);

    useEffect(() => {
        try {
            const savedSettings = localStorage.getItem('appSettings');
            if (savedSettings) {
                const parsedSettings: AppSettings = JSON.parse(savedSettings);
                // Basic validation
                if (parsedSettings && typeof parsedSettings.approverConfig === 'object') {
                    setSettings(parsedSettings);
                }
            }
        } catch (error) {
            console.error("Failed to load settings from local storage:", error);
        }
    }, []);
    
    const handleCreateNew = () => {
        setCurrentOrder(null);
        setView('form');
    }

    const handleEdit = (idToEdit: number) => {
        const orderToEdit = changeOrders.find(co => co.id === idToEdit);
        if (orderToEdit) {
            setCurrentOrder(orderToEdit);
            setView('form');
        }
    }

    const handleCancel = () => {
        setView('table');
        setCurrentOrder(null);
        setSelectedOrderIds(new Set());
    }

    const handleSave = (orderFromForm: ChangeOrder) => {
        let newStatus: ChangeOrderStatus = orderFromForm.status;
        const { approvals } = orderFromForm;

        const isRejected = approvals.some(app => app.status === 'Rejected');
        const isFullyApproved = approvals.length > 0 && approvals.every(app => app.status === 'Approved');

        if (isRejected) {
            newStatus = 'Rejected';
        } else if (isFullyApproved) {
            if (orderFromForm.status !== 'In Progress' && orderFromForm.status !== 'Completed') {
                newStatus = 'Approved';
            }
        } else {
             if (orderFromForm.status !== 'In Progress' && orderFromForm.status !== 'Completed') {
                newStatus = 'Pending Approval';
             }
        }
        
        const orderToSave = { ...orderFromForm, status: newStatus };

        if (orderToSave.id === 0) { // New order
            const newCoId = Math.max(0, ...changeOrders.map(co => co.id)) + 1;
            const newOrder = { ...orderToSave, id: newCoId };
            setChangeOrders([...changeOrders, newOrder].sort((a,b) => b.id - a.id));
        } else { // Existing order
            setChangeOrders(changeOrders.map(co => co.id === orderToSave.id ? orderToSave : co));
        }
        setView('table');
        setCurrentOrder(null);
        setSelectedOrderIds(new Set());
    }

    const handleDeleteSelected = () => {
        if (selectedOrderIds.size === 0) {
            alert("No change orders selected for deletion.");
            return;
        }
        setIsDeleteConfirmOpen(true);
    };

    const executeDelete = () => {
        setChangeOrders(prevOrders => prevOrders.filter(co => !selectedOrderIds.has(co.id)));
        setSelectedOrderIds(new Set());
        setIsDeleteConfirmOpen(false);
    };

    const handleRenumber = () => {
        if (changeOrders.length === 0) {
            alert("No change orders to renumber.");
            setIsRenumberConfirmOpen(false);
            return;
        }
        
        // Sort by date requested, oldest first
        const sortedOrders = [...changeOrders].sort((a, b) => new Date(a.dateRequested).getTime() - new Date(b.dateRequested).getTime());
        
        // Renumber based on sorted order
        const renumberedOrders = sortedOrders.map((order, index) => ({
            ...order,
            id: index + 1
        }));
        
        setChangeOrders(renumberedOrders);
        setSelectedOrderIds(new Set()); // Clear selections as IDs have changed
        setIsRenumberConfirmOpen(false);
        alert(`Successfully renumbered ${renumberedOrders.length} change orders.`);
    };


    const handleExportProject = () => {
        if (changeOrders.length === 0 && !settings.projectName) {
            alert("There is no project data to export.");
            return;
        }

        try {
            const projectData = { settings, changeOrders };
            const jsonData = JSON.stringify(projectData, null, 2);
            const blob = new Blob([jsonData], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `change-order-project-${settings.projectName || 'untitled'}.json`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            console.log('Project exported successfully.');
        } catch (error) {
            console.error('[Export Project] Failed to export project:', error);
            alert('Could not export the project.');
        }
    };

    const handleTriggerImport = () => {
        if (changeOrders.length > 0 && !confirm('Importing a new project will overwrite any unsaved changes. Do you want to continue?')) {
            return;
        }
        fileInputRef.current?.click();
    };

    // FIX: Changed preact.JSX to JSX after importing it.
    const handleFileImport = (event: JSX.TargetedEvent<HTMLInputElement, Event>) => {
        const target = event.target as HTMLInputElement;
        const file = target.files?.[0];
        if (!file) {
            return;
        }

        const reader = new FileReader();

        reader.onload = (e) => {
            const fileContent = e.target?.result as string;
            try {
                const data = JSON.parse(fileContent);
                if (!data.settings || !Array.isArray(data.changeOrders)) {
                    throw new Error("Invalid or corrupted project file structure.");
                }
                
                setSettings(data.settings);
                setChangeOrders(data.changeOrders);
                setCurrentOrder(null);
                setView('table');
                setSelectedOrderIds(new Set());
                alert('Project imported successfully!');

            } catch (error) {
                console.error('Failed to parse JSON file:', error);
                alert('Error: The selected file is not a valid project file.');
            } finally {
                // IMPORTANT: Reset the input value using the ref for robustness, ensuring the same file can be re-imported.
                if (fileInputRef.current) {
                    fileInputRef.current.value = '';
                }
            }
        };

        reader.onerror = () => {
            console.error("FileReader error:", reader.error);
            alert("An error occurred while reading the file.");
            // Also reset on error to allow retrying.
            if (fileInputRef.current) {
                fileInputRef.current.value = '';
            }
        };

        reader.readAsText(file);
    };

    const handleSaveSettings = (newSettings: AppSettings) => {
      try {
        localStorage.setItem('appSettings', JSON.stringify(newSettings));
        setSettings(newSettings);
        setIsSettingsOpen(false);
        alert("Settings saved.");
      } catch (error) {
        console.error("Failed to save settings:", error);
        alert("Could not save settings.");
      }
    };
    
    const handleEnterAdminMode = () => {
        setIsPasswordModalOpen(true);
    };

    const handleAdminLogin = (password: string) => {
        if (password === "0665") {
            setIsAdminMode(true);
            setIsPasswordModalOpen(false);
        } else {
            alert("Incorrect password.");
        }
    };


    return html`
        <header class="app-header">
            <span>Change Order Tracker ${settings.projectName ? `- ${settings.projectName}` : ''}</span>
            <span class="app-subtitle">Written by Tim Bickford 09/21/2025 Rev-1.0</span>
        </header>
        <main class="main-container">
            ${
                view === 'table' ? html`
                    <${ChangeOrderTable} 
                        changeOrders=${changeOrders} 
                        projectName=${settings.projectName}
                        projectLocation=${settings.projectLocation}
                        projectManager=${settings.projectManager}
                        selectedOrderIds=${selectedOrderIds}
                        setSelectedOrderIds=${setSelectedOrderIds}
                        onCreate=${handleCreateNew}
                        onEdit=${handleEdit}
                        onDeleteSelected=${handleDeleteSelected}
                        onExportProject=${handleExportProject}
                        onImportProject=${handleTriggerImport}
                        onConfigure=${() => setIsSettingsOpen(true)}
                        isAdminMode=${isAdminMode}
                        onEnterAdminMode=${handleEnterAdminMode}
                        onExitAdminMode=${() => setIsAdminMode(false)}
                        onRenumber=${() => setIsRenumberConfirmOpen(true)}
                    />` :
                view === 'form' ? html`
                    <div class="content-pane">
                      <${ChangeOrderForm} 
                        order=${currentOrder} 
                        approverConfig=${settings.approverConfig}
                        onSave=${handleSave} 
                        onCancel=${handleCancel} />
                    </div>` :
                null
            }
        </main>
        <input
            type="file"
            ref=${fileInputRef}
            onChange=${handleFileImport}
            style=${{ display: 'none' }}
            accept=".json,application/json"
        />
        ${isSettingsOpen && html`
            <${SettingsModal}
                settings=${settings}
                onSave=${handleSaveSettings}
                onClose=${() => setIsSettingsOpen(false)}
            />
        `}
        ${isDeleteConfirmOpen && html`
            <${ConfirmModal}
                isOpen=${isDeleteConfirmOpen}
                onClose=${() => setIsDeleteConfirmOpen(false)}
                onConfirm=${executeDelete}
                title="Confirm Deletion"
                confirmText="Confirm Delete"
                confirmClass="btn-danger"
            >
                <p>Are you sure you want to permanently delete ${selectedOrderIds.size} change order(s)? This action cannot be undone.</p>
            <//>
        `}
        ${isRenumberConfirmOpen && html`
            <${ConfirmModal}
                isOpen=${isRenumberConfirmOpen}
                onClose=${() => setIsRenumberConfirmOpen(false)}
                onConfirm=${handleRenumber}
                title="Confirm Mass Renumbering"
                confirmText="Yes, Renumber All"
                confirmClass="btn-danger"
            >
                <p>Are you sure you want to renumber all change orders? This will sort all orders by date requested (oldest first) and reassign their CO #s starting from 1. This action cannot be undone.</p>
            <//>
        `}
        ${isPasswordModalOpen && html`
            <${PasswordModal}
                isOpen=${isPasswordModalOpen}
                onClose=${() => setIsPasswordModalOpen(false)}
                onConfirm=${handleAdminLogin}
                title="Enter Admin Password"
            />
        `}
    `;
};

const SettingsModal = ({ settings, onSave, onClose }) => {
    const [localSettings, setLocalSettings] = useState(settings);

    const handleChange = (field: keyof AppSettings, value: string) => {
        setLocalSettings(prev => ({ ...prev, [field]: value }));
    };

    const handleApproverChange = (role: string, name: string) => {
        setLocalSettings(prev => ({
            ...prev,
            approverConfig: {
                ...prev.approverConfig,
                [role]: name
            }
        }));
    };

    const handleSave = () => {
        onSave(localSettings);
    };

    return html`
        <div class="modal-overlay" onClick=${onClose}>
            <div class="modal-content" onClick=${e => e.stopPropagation()}>
                <div class="modal-header">
                    <h2>Project Settings</h2>
                    <button class="btn-close" onClick=${onClose}>×</button>
                </div>
                <div class="modal-body">
                    <p>Define project details and approver names. These will be used across all change orders and reports.</p>
                    
                    <h3 class="form-section-title">Project Details</h3>
                    <div class="form-group">
                        <label for="projectName">Project Name</label>
                        <input id="projectName" type="text" class="form-input" value=${localSettings.projectName} onInput=${e => handleChange('projectName', e.target.value)} />
                    </div>
                     <div class="form-group">
                        <label for="projectLocation">Project Location</label>
                        <input id="projectLocation" type="text" class="form-input" value=${localSettings.projectLocation} onInput=${e => handleChange('projectLocation', e.target.value)} />
                    </div>
                    <div class="form-group">
                        <label for="projectManager">Project Manager Name</label>
                        <input id="projectManager" type="text" class="form-input" value=${localSettings.projectManager || ''} onInput=${e => handleChange('projectManager', e.target.value)} />
                    </div>

                    <h3 class="form-section-title">Approver Names</h3>
                    ${APPROVER_ROLES.map(role => html`
                        <div class="form-group" key=${role}>
                            <label for=${`approver-${role}`}>${role}</label>
                            <input
                                id=${`approver-${role}`}
                                type="text"
                                class="form-input"
                                value=${localSettings.approverConfig[role] || ''}
                                onInput=${e => handleApproverChange(role, e.target.value)}
                            />
                        </div>
                    `)}
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" onClick=${onClose}>Cancel</button>
                    <button type="button" class="btn btn-primary" onClick=${handleSave}>Save Settings</button>
                </div>
            </div>
        </div>
    `;
};

const ChangeOrderTable = ({ 
    changeOrders, projectName, projectLocation, projectManager, onEdit, onCreate, 
    onDeleteSelected, onExportProject, onImportProject, onConfigure, 
    selectedOrderIds, setSelectedOrderIds, isAdminMode, onEnterAdminMode, onExitAdminMode, onRenumber 
}) => {
    
    const [filters, setFilters] = useState({
        id: '',
        title: '',
        status: 'All',
        dateRequested: '',
        costMin: '',
        costMax: '',
        scheduleMin: '',
        scheduleMax: '',
    });

    const handleFilterChange = (e) => {
        const { name, value } = e.target;
        setFilters(prev => ({ ...prev, [name]: value }));
    };

    const clearFilters = () => {
        setFilters({
            id: '',
            title: '',
            status: 'All',
            dateRequested: '',
            costMin: '',
            costMax: '',
            scheduleMin: '',
            scheduleMax: '',
        });
    };
    
    const filteredChangeOrders = useMemo(() => {
        return changeOrders.filter(co => {
            const totalCost = co.costImpactEquipment + co.costImpactInstallation + co.costImpactOther;
            
            const idMatch = filters.id ? co.id.toString().includes(filters.id) : true;
            const titleMatch = filters.title ? co.title.toLowerCase().includes(filters.title.toLowerCase()) : true;
            const statusMatch = filters.status !== 'All' ? co.status === filters.status : true;
            const dateMatch = filters.dateRequested ? new Date(co.dateRequested) >= new Date(filters.dateRequested) : true;
            
            const costMin = parseFloat(filters.costMin);
            const costMax = parseFloat(filters.costMax);
            const costMinMatch = !isNaN(costMin) ? totalCost >= costMin : true;
            const costMaxMatch = !isNaN(costMax) ? totalCost <= costMax : true;

            const scheduleMin = parseInt(filters.scheduleMin, 10);
            const scheduleMax = parseInt(filters.scheduleMax, 10);
            const scheduleMinMatch = !isNaN(scheduleMin) ? co.scheduleImpactDays >= scheduleMin : true;
            const scheduleMaxMatch = !isNaN(scheduleMax) ? co.scheduleImpactDays <= scheduleMax : true;

            return idMatch && titleMatch && statusMatch && dateMatch && costMinMatch && costMaxMatch && scheduleMinMatch && scheduleMaxMatch;
        });
    }, [changeOrders, filters]);
    
    const handleGenerateReport = () => {
        if (changeOrders.length === 0) {
            alert("There are no change orders to report.");
            return;
        }
        generatePdfReport(changeOrders, projectName, projectLocation, projectManager);
    };

    const handleSelectAll = (e: JSX.TargetedEvent<HTMLInputElement, Event>) => {
        const isChecked = (e.target as HTMLInputElement).checked;
        setSelectedOrderIds(prevSelected => {
            const newSelected = new Set(prevSelected);
            if (isChecked) {
                filteredChangeOrders.forEach(co => newSelected.add(co.id));
            } else {
                filteredChangeOrders.forEach(co => newSelected.delete(co.id));
            }
            return newSelected;
        });
    };

    // FIX: Changed preact.JSX to JSX after importing it.
    const handleSelectSingle = (e: JSX.TargetedEvent<HTMLInputElement, Event>, id: number) => {
        const isChecked = (e.target as HTMLInputElement).checked;
        setSelectedOrderIds(prev => {
            const newSet = new Set(prev);
            if (isChecked) {
                newSet.add(id);
            } else {
                newSet.delete(id);
            }
            return newSet;
        });
    };

    const allFilteredIds = useMemo(() => new Set(filteredChangeOrders.map(co => co.id)), [filteredChangeOrders]);
    
    const isSelectAllChecked = useMemo(() => {
        if (allFilteredIds.size === 0) return false;
        for (const id of allFilteredIds) {
            if (!selectedOrderIds.has(id)) {
                return false;
            }
        }
        return true;
    }, [selectedOrderIds, allFilteredIds]);

    const totalCostImpact = useMemo(() => {
        return filteredChangeOrders.reduce((total, order) => {
            return total + order.costImpactEquipment + order.costImpactInstallation + order.costImpactOther;
        }, 0);
    }, [filteredChangeOrders]);


    return html`
        <div class="table-view">
            <header class="table-header">
                <h2>All Change Orders</h2>
                <div class="table-actions">
                     <button type="button" class="btn btn-danger" onClick=${onDeleteSelected} disabled=${selectedOrderIds.size === 0}>
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                            <path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"/>
                            <path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4zM2.5 3h11V2h-11z"/>
                        </svg>
                        Delete Selected (${selectedOrderIds.size})
                    </button>
                    <button type="button" class="btn btn-secondary" onClick=${onConfigure}>
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                            <path d="M9.405 1.05c-.413-1.4-2.397-1.4-2.81 0l-.1.34a1.464 1.464 0 0 1-2.105.872l-.31-.17c-1.283-.698-2.686.705-1.987 1.987l.169.311a1.464 1.464 0 0 1-.872 2.105l-.34.1c-1.4.413-1.4 2.397 0 2.81l.34.1a1.464 1.464 0 0 1 .872 2.105l-.17.31c-.698 1.283.705 2.686 1.987 1.987l.311-.169a1.464 1.464 0 0 1 2.105.872l.1.34c.413 1.4 2.397 1.4 2.81 0l.1-.34a1.464 1.464 0 0 1 2.105-.872l.31.17c1.283.698 2.686-.705 1.987-1.987l-.169-.311a1.464 1.464 0 0 1 .872-2.105l.34-.1c-1.4-.413-1.4-2.397 0-2.81l-.34-.1a1.464 1.464 0 0 1-.872-2.105l.17-.31c.698-1.283-.705-2.686-1.987-1.987l-.311.169a1.464 1.464 0 0 1-2.105-.872l-.1-.34zM8 10.93a2.929 2.929 0 1 1 0-5.86 2.929 2.929 0 0 1 0 5.858z"/>
                        </svg>
                        Settings
                    </button>
                    ${isAdminMode ? html`
                        <button type="button" class="btn btn-warning" onClick=${onRenumber}>Renumber All COs</button>
                        <button type="button" class="btn btn-secondary" onClick=${onExitAdminMode}>Exit Admin Mode</button>
                    ` : html `
                        <button type="button" class="btn btn-secondary" onClick=${onEnterAdminMode}>Admin Actions</button>
                    `}
                    <button type="button" class="btn btn-secondary" onClick=${onExportProject}>Export Project</button>
                    <button type="button" class="btn btn-secondary" onClick=${onImportProject}>Import Project</button>
                    <button type="button" class="btn btn-secondary" onClick=${handleGenerateReport}>
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                            <path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5z"/>
                            <path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708l3 3z"/>
                        </svg>
                        Generate PDF Report
                    </button>
                    <button type="button" class="btn btn-primary" onClick=${onCreate}>+ Add New Change Order</button>
                </div>
            </header>
            <div class="filter-bar">
                <div class="filter-group">
                    <input type="text" name="id" class="filter-input" placeholder="CO #" value=${filters.id} onInput=${handleFilterChange} />
                    <input type="text" name="title" class="filter-input" placeholder="Title contains..." value=${filters.title} onInput=${handleFilterChange} />
                    <select name="status" class="filter-select" value=${filters.status} onChange=${handleFilterChange}>
                        <option value="All">All Statuses</option>
                        <option>Pending Approval</option>
                        <option>Approved</option>
                        <option>Rejected</option>
                        <option>In Progress</option>
                        <option>Completed</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>Date after:</label>
                    <input type="date" name="dateRequested" class="filter-input" value=${filters.dateRequested} onInput=${handleFilterChange} />
                </div>
                <div class="filter-group">
                    <label>Cost:</label>
                    <input type="number" name="costMin" class="filter-input" placeholder="Min" value=${filters.costMin} onInput=${handleFilterChange} />
                    <span>-</span>
                    <input type="number" name="costMax" class="filter-input" placeholder="Max" value=${filters.costMax} onInput=${handleFilterChange} />
                </div>
                <div class="filter-group">
                    <label>Schedule (days):</label>
                    <input type="number" name="scheduleMin" class="filter-input" placeholder="Min" value=${filters.scheduleMin} onInput=${handleFilterChange} />
                    <span>-</span>
                    <input type="number" name="scheduleMax" class="filter-input" placeholder="Max" value=${filters.scheduleMax} onInput=${handleFilterChange} />
                </div>
                <button type="button" class="btn btn-secondary btn-sm" onClick=${clearFilters}>Clear Filters</button>
            </div>
            <div class="table-container">
                <table class="change-order-table">
                    <thead>
                        <tr>
                            <th>
                                <input
                                    type="checkbox"
                                    checked=${isSelectAllChecked}
                                    onChange=${handleSelectAll}
                                    aria-label="Select all change orders"
                                />
                            </th>
                            <th>CO #</th>
                            <th>Title</th>
                            <th>Status</th>
                            <th>Date Requested</th>
                            <th>Total Cost Impact</th>
                            <th>Schedule Impact</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${filteredChangeOrders.map(co => {
                            const totalCost = co.costImpactEquipment + co.costImpactInstallation + co.costImpactOther;
                            return html`
                                <tr key=${co.id} class=${selectedOrderIds.has(co.id) ? 'selected-row' : ''}>
                                    <td>
                                        <input
                                            type="checkbox"
                                            checked=${selectedOrderIds.has(co.id)}
                                            onChange=${(e) => handleSelectSingle(e, co.id)}
                                            aria-label=${`Select change order ${co.id}`}
                                        />
                                    </td>
                                    <td>${co.id}</td>
                                    <td>${co.title}</td>
                                    <td><span class="status-badge" style=${{backgroundColor: getStatusColor(co.status)}}>${co.status}</span></td>
                                    <td>${formatDate(co.dateRequested)}</td>
                                    <td>${formatCurrency(totalCost)}</td>
                                    <td>${co.scheduleImpactDays} days</td>
                                    <td class="action-cell">
                                        <button type="button" class="btn btn-secondary btn-sm" onClick=${() => onEdit(co.id)}>Edit</button>
                                    </td>
                                </tr>
                            `;
                        })}
                    </tbody>
                     ${filteredChangeOrders.length > 0 && html`
                        <tfoot>
                            <tr>
                                <td colSpan="5" class="total-label">Total (Filtered)</td>
                                <td class="total-value">${formatCurrency(totalCostImpact)}</td>
                                <td colSpan="2"></td>
                            </tr>
                        </tfoot>
                    `}
                </table>
                 ${filteredChangeOrders.length === 0 && changeOrders.length > 0 && html`
                    <div class="placeholder-table">
                        <p>No change orders match the current filters.</p>
                    </div>
                 `}
                 ${changeOrders.length === 0 && html`
                    <div class="placeholder-table">
                        <p>No change orders found. Click "+ Add New Change Order" to get started or import a project file.</p>
                    </div>
                `}
            </div>
        </div>
    `;
};

const emptyFormState: ChangeOrder = {
    id: 0, title: '', description: '', reason: '', status: 'Pending Approval',
    dateRequested: new Date().toISOString().split('T')[0], costImpactEquipment: 0,
    costImpactInstallation: 0, costImpactOther: 0, otherCostsExplanation: '', 
    scheduleImpactDays: 0, approvals: []
};

interface InternalApproverControl {
    role: string;
    included: boolean;
    status: ApprovalStatus;
    approvalDate: string;
}

const ChangeOrderForm = ({ order, onSave, onCancel, approverConfig }) => {
    const [formData, setFormData] = useState(order || emptyFormState);
    const [internalApprovers, setInternalApprovers] = useState<InternalApproverControl[]>([]);
    const [thirdPartyApprovers, setThirdPartyApprovers] = useState<Approval[]>([]);
    const [errors, setErrors] = useState<{ [key: string]: string }>({});

    const configuredInternalApprovers = useMemo(() => {
      return APPROVER_ROLES.filter(role => approverConfig[role]?.trim());
    }, [approverConfig]);

    useEffect(() => {
        const existingApprovals = order?.approvals || [];
        
        const newInternalState = configuredInternalApprovers.map(role => {
            const approverName = approverConfig[role];
            const existing = existingApprovals.find(app => app.name === approverName);
            return {
                role,
                included: !!existing,
                status: existing?.status || 'Pending',
                approvalDate: existing?.approvalDate || ''
            };
        });
        
        const newThirdPartyState = existingApprovals.filter(app => !Object.values(approverConfig).includes(app.name));

        setInternalApprovers(newInternalState);
        setThirdPartyApprovers(newThirdPartyState.map(a => ({...a, approvalDate: a.approvalDate || ''})));
        setFormData(order || emptyFormState);
        setErrors({}); // Clear errors on form load/reload
    }, [order, approverConfig, configuredInternalApprovers]);


    const isNew = !order;

    const handleChange = (e) => {
        const { name, value, type } = e.target;
        const isNumeric = type === 'number';

        // For numeric values, parseFloat and default to 0 if NaN (e.g., for empty string).
        const processedValue = isNumeric ? (parseFloat(value) || 0) : value;

        setFormData(prev => ({ ...prev, [name]: processedValue }));

        // Real-time validation for display: only check for negative numbers.
        // The raw `value` (string) from the event is used here.
        if (isNumeric) {
             if (parseFloat(value) < 0) { // isNaN check is implicit, parseFloat(invalid) < 0 is false
                setErrors(prev => ({...prev, [name]: 'Must be a non-negative number.'}));
            } else {
                const newErrors = {...errors};
                delete newErrors[name];
                setErrors(newErrors);
            }
        }
    }

    const handleInternalApproverChange = (index, field, value) => {
        const newApprovers = [...internalApprovers];
        const currentApprover = newApprovers[index];
        currentApprover[field] = value;

        // UX Improvement: If date is set and status is Pending, auto-approve.
        if (field === 'approvalDate' && value && currentApprover.status === 'Pending') {
            currentApprover.status = 'Approved';
        }

        // If status is changed to Pending, clear the date
        if (field === 'status' && value === 'Pending') {
            currentApprover.approvalDate = '';
        }
        // if unchecking, reset to default
        if (field === 'included' && !value) {
            currentApprover.status = 'Pending';
            currentApprover.approvalDate = '';
        }
        setInternalApprovers(newApprovers);
    };

    const handleThirdPartyApproverChange = (index, field, value) => {
        const newApprovals = [...thirdPartyApprovers];
        const currentApprover = newApprovals[index];
        currentApprover[field] = value;

        // UX Improvement: If date is set and status is Pending, auto-approve.
        if (field === 'approvalDate' && value && currentApprover.status === 'Pending') {
            currentApprover.status = 'Approved';
        }

         // If status is changed to Pending, clear the date
        if (field === 'status' && value === 'Pending') {
            currentApprover.approvalDate = '';
        }
        setThirdPartyApprovers(newApprovals);
    }

    const addThirdPartyApprover = () => {
        setThirdPartyApprovers(prev => [...prev, {name: '', status: 'Pending', approvalDate: ''}]);
    }

    const removeThirdPartyApprover = (index) => {
        setThirdPartyApprovers(prev => prev.filter((_, i) => i !== index));
    }

    const handleSubmit = (e) => {
        e.preventDefault();
        
        // Create a copy of existing errors to update.
        const newErrors = {...errors};

        // Clear previous errors for these specific fields before re-validating.
        delete newErrors.title;
        delete newErrors.description;
        delete newErrors.reason;

        // Re-validate the required fields and update the errors object.
        if (!formData.title.trim()) newErrors.title = 'Title is required.';
        if (!formData.description.trim()) newErrors.description = 'Description is required.';
        if (!formData.reason.trim()) newErrors.reason = 'Reason for change is required.';
        
        // Update the state to display all current errors.
        setErrors(newErrors);

        // Block submission ONLY if one of the required text fields is empty.
        // This allows saving even if other non-critical fields have validation errors.
        if (!formData.title.trim() || !formData.description.trim() || !formData.reason.trim()) {
            alert('Please fill out Title, Description, and Reason for Change before saving.');
            return;
        }

        const finalApprovals: Approval[] = [];
        
        internalApprovers.forEach(app => {
            if (app.included) {
                const approverName = approverConfig[app.role];
                finalApprovals.push({ name: approverName, status: app.status, approvalDate: app.approvalDate || null });
            }
        });

        thirdPartyApprovers.forEach(app => {
            if(app.name) { // only add if name is not empty
                finalApprovals.push({ ...app, approvalDate: app.approvalDate || null });
            }
        });
        
        onSave({...formData, approvals: finalApprovals});
    }

    const totalCost = useMemo(() => {
        return formData.costImpactEquipment + formData.costImpactInstallation + formData.costImpactOther;
    }, [formData.costImpactEquipment, formData.costImpactInstallation, formData.costImpactOther]);

    return html`
        <form class="form-view" onSubmit=${handleSubmit} noValidate>
            <h2>${isNew ? 'Create New Change Order' : `Editing #${order.id}`}</h2>
            <div class="form-grid">
                <div class="form-group full-width">
                    <label for="title">Title</label>
                    <input id="title" name="title" type="text" class=${`form-input ${errors.title ? 'is-invalid' : ''}`} required value=${formData.title} onInput=${handleChange} />
                    ${errors.title && html`<div class="error-message">${errors.title}</div>`}
                </div>
                <div class="form-group full-width">
                    <label for="description">Description</label>
                    <textarea id="description" name="description" class=${`form-textarea ${errors.description ? 'is-invalid' : ''}`} required value=${formData.description} onInput=${handleChange}></textarea>
                    ${errors.description && html`<div class="error-message">${errors.description}</div>`}
                </div>
                <div class="form-group full-width">
                    <label for="reason">Reason for Change</label>
                    <textarea id="reason" name="reason" class=${`form-textarea ${errors.reason ? 'is-invalid' : ''}`} required value=${formData.reason} onInput=${handleChange}></textarea>
                    ${errors.reason && html`<div class="error-message">${errors.reason}</div>`}
                </div>
                
                <div class="form-group">
                    <label for="dateRequested">Date Requested</label>
                    <input id="dateRequested" name="dateRequested" type="date" class="form-input" required value=${formData.dateRequested} onInput=${handleChange} />
                </div>
                <div class="form-group">
                    <label for="status">Status</label>
                    <select id="status" name="status" class="form-select" value=${formData.status} onChange=${handleChange}>
                        <option>Pending Approval</option>
                        <option>Approved</option>
                        <option>Rejected</option>
                        <option>In Progress</option>
                        <option>Completed</option>
                    </select>
                </div>

                <h3 class="form-section-title">Impact Analysis</h3>
                <div class="form-group">
                    <label for="costImpactEquipment">Equipment Cost</label>
                    <input id="costImpactEquipment" name="costImpactEquipment" type="number" step="0.01" min="0" class=${`form-input ${errors.costImpactEquipment ? 'is-invalid' : ''}`} value=${formData.costImpactEquipment || ''} onInput=${handleChange} />
                    ${errors.costImpactEquipment && html`<div class="error-message">${errors.costImpactEquipment}</div>`}
                </div>
                <div class="form-group">
                    <label for="costImpactInstallation">Installation Cost</label>
                    <input id="costImpactInstallation" name="costImpactInstallation" type="number" step="0.01" min="0" class=${`form-input ${errors.costImpactInstallation ? 'is-invalid' : ''}`} value=${formData.costImpactInstallation || ''} onInput=${handleChange} />
                     ${errors.costImpactInstallation && html`<div class="error-message">${errors.costImpactInstallation}</div>`}
                </div>
                <div class="form-group">
                    <label for="costImpactOther">Other Costs</label>
                    <input id="costImpactOther" name="costImpactOther" type="number" step="0.01" min="0" class=${`form-input ${errors.costImpactOther ? 'is-invalid' : ''}`} value=${formData.costImpactOther || ''} onInput=${handleChange} />
                    ${errors.costImpactOther && html`<div class="error-message">${errors.costImpactOther}</div>`}
                </div>
                 <div class="form-group">
                    <label for="totalCost">Total Cost Impact</label>
                    <input id="totalCost" name="totalCost" type="text" class="form-input" value=${formatCurrency(totalCost)} readOnly />
                </div>
                
                ${formData.costImpactOther > 0 && html`
                    <div class="form-group full-width">
                        <label for="otherCostsExplanation">Other Costs Explanation</label>
                        <textarea id="otherCostsExplanation" name="otherCostsExplanation" class="form-textarea" placeholder="Provide a breakdown or explanation for other costs..." value=${formData.otherCostsExplanation || ''} onInput=${handleChange}></textarea>
                    </div>
                `}

                <div class="form-group">
                    <label for="scheduleImpactDays">Schedule Impact (Days)</label>
                    <input id="scheduleImpactDays" name="scheduleImpactDays" type="number" min="0" class=${`form-input ${errors.scheduleImpactDays ? 'is-invalid' : ''}`} value=${formData.scheduleImpactDays || ''} onInput=${handleChange} />
                    ${errors.scheduleImpactDays && html`<div class="error-message">${errors.scheduleImpactDays}</div>`}
                </div>
                <div class="form-group full-width"></div>

                <h3 class="form-section-title">Internal Approvals</h3>
                ${configuredInternalApprovers.length > 0 ? html`
                    <div class="internal-approvals-header">
                        <span>Approver</span>
                        <span>Status</span>
                        <span>Approval Date</span>
                    </div>
                    ${internalApprovers.map((app, index) => html`
                        <div class="internal-approver-row" key=${app.role}>
                            <div class="approver-name">
                                <input type="checkbox" id=${`include-${index}`} checked=${app.included} onChange=${e => handleInternalApproverChange(index, 'included', e.target.checked)} />
                                <label for=${`include-${index}`}>
                                  ${approverConfig[app.role]}
                                  <span class="approver-role-label">(${app.role})</span>
                                </label>
                            </div>
                            <select class="form-select" value=${app.status} onChange=${e => handleInternalApproverChange(index, 'status', e.target.value)} disabled=${!app.included}>
                                <option>Pending</option>
                                <option>Approved</option>
                                <option>Rejected</option>
                            </select>
                            <input type="date" class="form-input" value=${app.approvalDate} onInput=${e => handleInternalApproverChange(index, 'approvalDate', e.target.value)} disabled=${!app.included || app.status === 'Pending'} />
                        </div>
                    `)}
                ` : html`
                  <div class="placeholder-form full-width">
                    <p>No internal approvers have been configured. Please add them in the settings to see them here.</p>
                  </div>
                `}

                <h3 class="form-section-title">Third-Party Approvals</h3>
                ${thirdPartyApprovers.map((app, index) => html`
                    <div class="approver-row" key=${index}>
                        <div class="form-group">
                           ${index === 0 ? html`<label>Approver Name</label>` : ''}
                           <input type="text" class="form-input" value=${app.name} onInput=${e => handleThirdPartyApproverChange(index, 'name', e.target.value)} placeholder="e.g. External Regulator" />
                        </div>
                        <div class="form-group">
                            ${index === 0 ? html`<label>Status</label>` : ''}
                            <select class="form-select" value=${app.status} onChange=${e => handleThirdPartyApproverChange(index, 'status', e.target.value)}>
                                <option>Pending</option>
                                <option>Approved</option>
                                <option>Rejected</option>
                            </select>
                        </div>
                        <div class="form-group">
                            ${index === 0 ? html`<label>Approval Date</label>` : ''}
                            <input type="date" class="form-input" value=${app.approvalDate} onInput=${e => handleThirdPartyApproverChange(index, 'approvalDate', e.target.value)} disabled=${app.status === 'Pending'} />
                        </div>
                        <button type="button" class="btn btn-danger btn-remove-approver" onClick=${() => removeThirdPartyApprover(index)}>Remove</button>
                    </div>
                `)}
                <div class="form-group full-width">
                    <button type="button" class="btn btn-secondary" onClick=${addThirdPartyApprover}>+ Add Third-Party Approver</button>
                </div>
                
                <div class="form-actions">
                    <button type="button" class="btn btn-secondary" onClick=${onCancel}>Cancel</button>
                    <button type="submit" class="btn btn-primary">Save Changes</button>
                </div>
            </div>
        </form>
    `;
};

render(html`<${App} />`, document.getElementById("root"));