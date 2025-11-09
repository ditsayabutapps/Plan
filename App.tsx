import React, { useState, useMemo, useCallback, useRef } from 'react';
import { Project } from './types';
import { Header } from './components/Header';
import { SummaryBar } from './components/SummaryBar';
import { ProjectTable } from './components/ProjectTable';
import { PlusIcon } from './components/icons/PlusIcon';
import { ImportIcon } from './components/icons/ImportIcon';
import { ExportIcon } from './components/icons/ExportIcon';


declare const XLSX: any;

const App: React.FC = () => {
    const [fiscalYear, setFiscalYear] = useState(new Date().getFullYear() + 543);
    const [projects, setProjects] = useState<Project[]>([
        {
            id: crypto.randomUUID(),
            name: 'จัดทำเว็บไซต์หน่วยงาน',
            department: 'ฝ่ายเทคโนโลยีสารสนเทศ',
            amount: '150000',
            manager: 'นายสมชาย ใจดี',
            duration: 'ม.ค. - มี.ค.'
        },
        {
            id: crypto.randomUUID(),
            name: 'โครงการอบรมพนักงาน',
            department: 'ฝ่ายบุคคล',
            amount: '85000',
            manager: 'นางสาวสมศรี มีสุข',
            duration: 'เม.ย.'
        },
        {
            id: crypto.randomUUID(),
            name: '',
            department: '',
            amount: '',
            manager: '',
            duration: ''
        }
    ]);
    const [totalBudget, setTotalBudget] = useState(1000000);
    const fileInputRef = useRef<HTMLInputElement>(null);

    const handleBudgetChange = useCallback((newBudget: number) => {
        setTotalBudget(isNaN(newBudget) ? 0 : newBudget);
    }, []);

    const handleInputChange = useCallback((index: number, field: keyof Omit<Project, 'id'>, value: string) => {
        setProjects(currentProjects => {
            const newProjects = [...currentProjects];
            newProjects[index] = { ...newProjects[index], [field]: value };
            return newProjects;
        });
    }, []);

    const handleAddRow = useCallback(() => {
        setProjects(currentProjects => [
            ...currentProjects,
            {
                id: crypto.randomUUID(),
                name: '',
                department: '',
                amount: '',
                manager: '',
                duration: ''
            }
        ]);
    }, []);

    const handleDeleteRow = useCallback((index: number) => {
        setProjects(currentProjects => currentProjects.filter((_, i) => i !== index));
    }, []);

    const handleExport = useCallback(() => {
        const thaiHeaders = {
            name: 'งาน/โครงการ',
            department: 'ฝ่ายงานรับผิดชอบ',
            amount: 'จำนวนเงินที่เสนอขออนุมัติ',
            manager: 'ผู้รับผิดชอบโครงการ',
            duration: 'ระยะเวลาดำเนินงาน',
        };

        const dataToExport = projects.map((p, index) => ({
            'ลำดับที่': index + 1,
            [thaiHeaders.name]: p.name,
            [thaiHeaders.department]: p.department,
            [thaiHeaders.amount]: parseFloat(String(p.amount).replace(/,/g, '')) || 0,
            [thaiHeaders.manager]: p.manager,
            [thaiHeaders.duration]: p.duration,
        }));
        
        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Projects');
        XLSX.writeFile(workbook, `แผนงานและโครงการ-ปีงบประมาณ-${fiscalYear}.xlsx`);
    }, [projects, fiscalYear]);

    const handleImport = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target?.result as ArrayBuffer);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const json: any[] = XLSX.utils.sheet_to_json(worksheet);

                const englishToThaiMap = {
                    'name': 'งาน/โครงการ',
                    'department': 'ฝ่ายงานรับผิดชอบ',
                    'amount': 'จำนวนเงินที่เสนอขออนุมัติ',
                    'manager': 'ผู้รับผิดชอบโครงการ',
                    'duration': 'ระยะเวลาดำเนินงาน',
                };
                const thaiToEnglishMap = Object.fromEntries(Object.entries(englishToThaiMap).map(([k, v]) => [v, k]));


                const importedProjects: Project[] = json.map(row => {
                    const project: Partial<Project> & { id: string } = { id: crypto.randomUUID() };
                    for (const thaiHeader in thaiToEnglishMap) {
                        if (row[thaiHeader] !== undefined) {
                            const englishKey = thaiToEnglishMap[thaiHeader] as keyof Omit<Project, 'id'>;
                            project[englishKey] = String(row[thaiHeader]);
                        }
                    }
                    return project as Project;
                }).filter(p => p.name); // Filter out rows without a project name

                if (importedProjects.length > 0) {
                    setProjects(importedProjects);
                } else {
                    alert("ไม่พบข้อมูลที่ถูกต้องในไฟล์ Excel หรือไฟล์อาจจะว่างเปล่า");
                }
            } catch (error) {
                console.error("Error importing file:", error);
                alert("เกิดข้อผิดพลาดในการนำเข้าไฟล์ โปรดตรวจสอบว่าไฟล์เป็นรูปแบบ .xlsx ที่ถูกต้อง");
            } finally {
                // Reset file input to allow re-uploading the same file
                if (fileInputRef.current) {
                    fileInputRef.current.value = '';
                }
            }
        };
        reader.readAsArrayBuffer(file);
    }, []);

    const summary = useMemo(() => {
        const totalProjects = projects.filter(p => p.name.trim() !== '').length;
        const totalAmount = projects.reduce((sum, project) => {
            const amount = parseFloat(String(project.amount).replace(/,/g, ''));
            return sum + (isNaN(amount) ? 0 : amount);
        }, 0);
        const remainingBudget = totalBudget - totalAmount;
        return { totalProjects, totalAmount, remainingBudget };
    }, [projects, totalBudget]);

    return (
        <div className="min-h-screen bg-slate-50 text-slate-800 font-sans">
            <div className="container mx-auto p-4 sm:p-6 lg:p-8">
                <Header fiscalYear={fiscalYear} onFiscalYearChange={setFiscalYear} />
                <SummaryBar 
                    totalProjects={summary.totalProjects} 
                    totalAmount={summary.totalAmount}
                    totalBudget={totalBudget}
                    remainingBudget={summary.remainingBudget}
                    onBudgetChange={handleBudgetChange}
                />
                
                <main className="mt-8 bg-white shadow-lg rounded-xl overflow-hidden">
                     <ProjectTable 
                        projects={projects}
                        onInputChange={handleInputChange}
                        onDeleteRow={handleDeleteRow}
                     />
                     <div className="p-4 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center gap-4">
                        <button 
                            onClick={handleAddRow}
                            className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-indigo-600 rounded-lg hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 transition-all duration-200"
                        >
                            <PlusIcon className="w-5 h-5" />
                            เพิ่มแถว
                        </button>
                        <input
                            type="file"
                            ref={fileInputRef}
                            onChange={handleImport}
                            accept=".xlsx, .xls"
                            className="hidden"
                            id="file-upload"
                        />
                        <button 
                            onClick={() => fileInputRef.current?.click()}
                            className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-emerald-600 rounded-lg hover:bg-emerald-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-emerald-500 transition-all duration-200"
                        >
                            <ImportIcon className="w-5 h-5" />
                            นำเข้า Excel
                        </button>
                        <button 
                            onClick={handleExport}
                            className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-sky-600 rounded-lg hover:bg-sky-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-sky-500 transition-all duration-200"
                        >
                            <ExportIcon className="w-5 h-5" />
                            ส่งออก Excel
                        </button>
                    </div>
                </main>
            </div>
        </div>
    );
};

export default App;