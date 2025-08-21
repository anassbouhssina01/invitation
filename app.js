document.addEventListener('DOMContentLoaded', () => {
    console.log("app.js loaded");

    // --- Element Selectors ---
    const jsonUpload = document.getElementById('json-upload');
    const toggleJsonUploadBtn = document.getElementById('toggle-json-upload');
    const jsonUploadContainer = document.getElementById('json-upload-container');

    const excelUpload = document.getElementById('excel-upload');
    const toggleExcelUploadBtn = document.getElementById('toggle-excel-upload');
    const excelUploadContainer = document.getElementById('excel-upload-container');

    const searchInput = document.getElementById('search-input');
    const searchCityInput = document.getElementById('search-city');
    const searchPprInput = document.getElementById('search-ppr');
    const searchResults = document.getElementById('search-results');
    
    const invitedListDiv = document.getElementById('invited-list');
    const wordTemplateUpload = document.getElementById('word-template-upload');
    const generateWordBtn = document.getElementById('generate-word');

    // --- State Variables ---
    let employeeData = [];
    let invitedList = [];
    let wordTemplate = null;

    // --- UI Toggles ---
    toggleJsonUploadBtn.addEventListener('click', () => {
        jsonUploadContainer.classList.toggle('hidden');
    });

    toggleExcelUploadBtn.addEventListener('click', () => {
        excelUploadContainer.classList.toggle('hidden');
    });

    // --- Data Handling ---
    function loadEmployeeData() {
        const data = localStorage.getItem('employeeData');
        employeeData = data ? JSON.parse(data) : [];
    }

    function loadInvitedList() {
        const data = localStorage.getItem('invitedListData');
        invitedList = data ? JSON.parse(data) : [];
        renderInvitedList(); // Render list after loading
    }

    function saveInvitedList() {
        localStorage.setItem('invitedListData', JSON.stringify(invitedList));
    }

    // --- JSON Upload ---
    jsonUpload.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = JSON.parse(e.target.result);
                if (data.employees && Array.isArray(data.employees)) {
                    localStorage.setItem('employeeData', JSON.stringify(data.employees));
                    loadEmployeeData();
                    alert('تم تحميل بيانات الموظفين بنجاح!');
                    jsonUploadContainer.classList.add('hidden');
                } else {
                    alert('ملف JSON غير صالح. يجب أن يحتوي على مصفوفة باسم "employees".');
                }
            } catch (error) {
                console.error('Error parsing JSON:', error);
                alert('حدث خطأ أثناء قراءة الملف. الرجاء التأكد من أن الملف بصيغة JSON صحيحة.');
            }
        };
        reader.readAsText(file);
    });

    // --- Search Functionality ---
    function performSearch() {
        const searchTerm = searchInput.value.toLowerCase().trim();
        const cityTerm = searchCityInput.value.toLowerCase().trim();
        const pprTerm = searchPprInput.value.toLowerCase().trim();
        searchResults.innerHTML = '';

        if (searchTerm.length === 0 && cityTerm.length === 0 && pprTerm.length === 0) {
            return;
        }

        const filteredEmployees = employeeData.filter(emp => {
            const empName = (emp.fullName || '').toLowerCase();
            const empId = (emp.employeeId || '').toString();
            const empCity = (emp.city || '').toLowerCase();
            const empPpr = (emp.workLocation || '').toLowerCase();

            const nameMatch = searchTerm ? (empName.includes(searchTerm) || empId.includes(searchTerm)) : true;
            const cityMatch = cityTerm ? empCity.includes(cityTerm) : true;
            const pprMatch = pprTerm ? empPpr.includes(pprTerm) : true;

            return nameMatch && cityMatch && pprMatch;
        });

        displaySearchResults(filteredEmployees);
    }

    searchInput.addEventListener('keyup', performSearch);
    searchCityInput.addEventListener('keyup', performSearch);
    searchPprInput.addEventListener('keyup', performSearch);

    function displaySearchResults(employees) {
        searchResults.innerHTML = '';
        employees.forEach(emp => {
            const empDiv = document.createElement('div');
            empDiv.className = 'employee-item';
            empDiv.innerHTML = `
                <div class="employee-info">
                    <span>${emp.fullName} (الرقم: ${emp.employeeId})</span>
                    <span class="details">${emp.workLocation} | ${emp.city} | ${emp.division}</span>
                </div>
            `;
            empDiv.dataset.employeeId = emp.employeeId;
            
            empDiv.addEventListener('click', () => {
                addToInvitedList(emp.employeeId);
            });

            searchResults.appendChild(empDiv);
        });
    }

    // --- Invited List Management ---
    function renderInvitedList() {
        invitedListDiv.innerHTML = '';
        invitedList.forEach(emp => {
            const item = document.createElement('div');
            item.className = 'employee-item';
            item.innerHTML = `
                <div class="employee-info">
                    <span>${emp.fullName} (الرقم: ${emp.employeeId})</span>
                    <span class="details">${emp.workLocation} | ${emp.division} | ${emp.city}</span>
                </div>
                <button class="delete-btn" data-employee-id="${emp.employeeId}">حذف</button>
            `;
            invitedListDiv.appendChild(item);
        });
    }

    function addToInvitedList(employeeId) {
        // Find the employee in the main data list, converting ID to integer
        const employeeToAdd = employeeData.find(emp => parseInt(emp.employeeId, 10) === parseInt(employeeId, 10));
        
        if (employeeToAdd && !invitedList.some(emp => parseInt(emp.employeeId, 10) === parseInt(employeeId, 10))) {
            invitedList.push(employeeToAdd);
            saveInvitedList();
            renderInvitedList();
            alert(`تمت إضافة "${employeeToAdd.fullName}" إلى لائحة المدعوين.`);
            // Clear search inputs and results
            searchInput.value = '';
            searchCityInput.value = '';
            searchPprInput.value = '';
            searchResults.innerHTML = '';
        } else if (!employeeToAdd) {
             alert('لم يتم العثور على الموظف.');
        } else {
            alert('هذا الموظف موجود بالفعل في القائمة.');
        }
    }

    function removeFromInvitedList(employeeId) {
        invitedList = invitedList.filter(emp => parseInt(emp.employeeId, 10) !== parseInt(employeeId, 10));
        saveInvitedList();
        renderInvitedList();
    }

    invitedListDiv.addEventListener('click', (event) => {
        if (event.target.classList.contains('delete-btn')) {
            const employeeId = event.target.dataset.employeeId;
            removeFromInvitedList(employeeId);
        }
    });

    // --- Excel Import ---
    excelUpload.addEventListener('change', (event) => {
        if (typeof XLSX === 'undefined') {
            alert('عذرًا، حدث خطأ أثناء تحميل مكتبة معالجة ملفات Excel.');
            return;
        }

        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet);
                
                let importCount = 0;
                json.forEach(row => {
                    const employeeId = row.employeeId;
                    if (employeeId) {
                        const employeeToAdd = employeeData.find(emp => parseInt(emp.employeeId, 10) === parseInt(employeeId, 10));
                        if (employeeToAdd && !invitedList.some(emp => parseInt(emp.employeeId, 10) === parseInt(employeeId, 10))) {
                             addToInvitedList(employeeId);
                             importCount++;
                        }
                    }
                });
                
                if (importCount > 0) {
                    alert(`تم استيراد وإضافة ${importCount} موظف بنجاح.`);
                } else {
                    alert('لم يتم العثور على موظفين جدد في الملف لإضافتهم.');
                }
            } catch (error) {
                console.error("Error processing Excel file:", error);
                alert("حدث خطأ أثناء معالجة ملف Excel.");
            } finally {
                excelUpload.value = '';
                excelUploadContainer.classList.add('hidden');
            }
        };
        reader.readAsArrayBuffer(file);
    });

    // --- Word Document Generation ---
    wordTemplateUpload.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) {
            wordTemplate = null;
            return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
            wordTemplate = e.target.result;
            alert('تم تحميل قالب Word بنجاح.');
        };
        reader.readAsArrayBuffer(file);
    });

    function processGroupData(employees) {
        const males = employees.filter(e => e.gender === 'السيد');
        const females = employees.filter(e => e.gender === 'السيدة');
        
        let maleCollectiveTitle = '';
        if (males.length > 0) {
            const title = males.length === 1 ? 'السيد' : (males.length === 2 ? 'السيدان' : 'السادة');
            const maleNames = males.map(e => e.fullName).join(' و ');
            maleCollectiveTitle = `${title} ${maleNames}`;
        }

        let femaleCollectiveTitle = '';
        if (females.length > 0) {
            const title = females.length === 1 ? 'السيدة' : (females.length === 2 ? 'السيدتان' : 'السيدات');
            const femaleNames = females.map(e => e.fullName).join(' و ');
            femaleCollectiveTitle = `${title} ${femaleNames}`;
        }

        if (males.length > 0 && females.length > 0) {
            femaleCollectiveTitle = `و ${femaleCollectiveTitle}`;
        }

        let computedVar = '';
        const totalCount = employees.length;
        if (totalCount === 1) {
            computedVar = males.length === 1 ? 'المعني' : 'المعنية';
        } else if (totalCount === 2) {
            if (males.length === 2) computedVar = 'المعنيان';
            else if (females.length === 2) computedVar = 'المعنيتان';
            else computedVar = 'المعنيان';
        } else {
             computedVar = males.length > 0 ? 'المعنيون' : 'المعنيات';
        }

        const city = employees.length > 0 ? employees[0].city : '';

        const jobTitleGrammar = {
            'وكيل': { m: {s: 'وكيل', d: 'وكيلان', p: 'وكلاء'}, f: {s: 'وكيلة', d: 'وكيلتان', p: 'وكيلات'} },
            'قاض': { m: {s: 'قاض', d: 'قاضيان', p: 'قضاة'}, f: {s: 'قاضية', d: 'قاضيتان', p: 'قاضيات'} },
            'قاضي': { m: {s: 'قاضي', d: 'قاضيان', p: 'قضاة'}, f: {s: 'قاضية', d: 'قاضيتان', p: 'قاضيات'} },
            'مستشار': { m: {s: 'مستشار', d: 'مستشاران', p: 'مستشارون'}, f: {s: 'مستشارة', d: 'مستشارتان', p: 'مستشارات'} },
            'وكيل عام': { m: {s: 'وكيل عام', d: 'وكيلا عام', p: 'وكلاء عامون'}, f: {s: 'وكيلة عامة', d: 'وكيلتان عامتان', p: 'وكيلات عامات'} },
            'محام عام': { m: {s: 'محام عام', d: 'محاميان عامان', p: 'محامون عامون'}, f: {s: 'محامية عامة', d: 'محاميتان عامتان', p: 'محاميات عامات'} },
            'نائب الوكيل العام للملك': { m: {s: 'نائب الوكيل العام للملك', d: 'نائبا الوكيل العام للملك', p: 'نواب الوكيل العام للملك'}, f: {s: 'نائبة الوكيل العام للملك', d: 'نائبتا الوكيل العام للملك', p: 'نائبات الوكيل العام للملك'} },
            'وكيل الملك': { m: {s: 'وكيل الملك', d: 'وكيلا الملك', p: 'وكلاء الملك'}, f: {s: 'وكيلة الملك', d: 'وكيلتا الملك', p: 'وكيلات الملك'} },
            'نائب وكيل الملك': { m: {s: 'نائب وكيل الملك', d: 'نائبا وكيل الملك', p: 'نواب وكيل الملك'}, f: {s: 'نائبة وكيل الملك', d: 'نائبتا وكيل الملك', p: 'نائبات وكيل الملك'} },
        };

        let combinedJobTitle = '';
        if (employees.length > 0) {
            const rootCounts = employees.reduce((acc, e) => {
                const title = (e.jobTitle || '').trim();
                if (title) {
                    const root = title.replace(/ة$/, '');
                    acc[root] = (acc[root] || 0) + 1;
                }
                return acc;
            }, {});

            let dominantRoot = '';
            let maxCount = 0;
            for (const root in rootCounts) {
                if (rootCounts[root] > maxCount) {
                    maxCount = rootCounts[root];
                    dominantRoot = root;
                }
            }

            if (jobTitleGrammar[dominantRoot]) {
                const grammarSet = jobTitleGrammar[dominantRoot];
                const relevantEmployees = employees.filter(e => (e.jobTitle || '').trim().replace(/ة$/, '') === dominantRoot);
                const groupMaleCount = relevantEmployees.filter(e => e.gender === 'السيد').length;
                const groupFemaleCount = relevantEmployees.filter(e => e.gender === 'السيدة').length;
                const groupTotalCount = relevantEmployees.length;

                if (groupTotalCount === 1) {
                    combinedJobTitle = groupMaleCount === 1 ? grammarSet.m.s : grammarSet.f.s;
                } else if (groupTotalCount === 2) {
                    if (groupMaleCount === 2) combinedJobTitle = grammarSet.m.d;
                    else if (groupFemaleCount === 2) combinedJobTitle = grammarSet.f.d;
                    else combinedJobTitle = grammarSet.m.d;
                } else if (groupTotalCount > 2){
                    combinedJobTitle = groupMaleCount > 0 ? grammarSet.m.p : grammarSet.f.p;
                } else {
                    combinedJobTitle = dominantRoot || '';
                }
            } else {
                combinedJobTitle = dominantRoot || '';
            }
        }
        
        return { city, maleCollectiveTitle, femaleCollectiveTitle, computedVar, combinedJobTitle };
    }

    generateWordBtn.addEventListener('click', () => {
        try {
            if (!wordTemplate) {
                alert('الرجاء تحميل قالب Word أولاً.');
                return;
            }
            if (invitedList.length === 0) {
                alert('لائحة المدعوين فارغة. الرجاء إضافة موظفين أولاً.');
                return;
            }
            if (typeof PizZip === 'undefined' || typeof docxtemplater === 'undefined') {
                alert('عذرًا، حدث خطأ أثناء تحميل مكتبات إنشاء المستندات.');
                return;
            }

            const groupedEmployees = invitedList.reduce((acc, emp) => {
                const key = `${emp.workLocation || 'N/A'}_${emp.division || 'N/A'}`;
                if (!acc[key]) {
                    acc[key] = {
                        workLocation: emp.workLocation,
                        division: emp.division,
                        employees: []
                    };
                }
                acc[key].employees.push(emp);
                return acc;
            }, {});

            for (const key in groupedEmployees) {
                const group = groupedEmployees[key];
                const { workLocation, division, employees } = group;

                const responsiblePerson = employeeData.find(e => {
                    // Check if workLocation, division, and city match
                    if (e.workLocation !== workLocation || e.division !== division || e.city !== employees[0].city) return false;
                    
                    // Check if postResponsibility exists and is a main title
                    if (!e.postResponsibility) return false;
                    const title = e.postResponsibility.trim();
                    return (title.includes("رئيس") || title.includes("رئيسة") || title.includes("وكيل") || title.includes("وكيلة"));
                });

                let responsibilityLine = '';
                if (responsiblePerson) {
                    const genderPrefix = responsiblePerson.gender === 'السيدة' ? 'السيدة' : 'السيد';
                    responsibilityLine = `${responsiblePerson.workLocation} ${genderPrefix} ${responsiblePerson.postResponsibility}`;
                }
                
                const responsiblePersonId = responsiblePerson ? responsiblePerson.employeeId : null;
                const filteredEmployees = employees.filter(e => e.employeeId !== responsiblePersonId);
                const processedData = processGroupData(filteredEmployees);

                const zip = new PizZip(wordTemplate);
                const doc = new docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

                doc.setData({
                    workLocation: workLocation,
                    division: division,
                    employees: filteredEmployees,
                    responsibilityLine: responsibilityLine,
                    ...processedData
                });

                doc.render();

                const out = doc.getZip().generate({
                    type: 'blob',
                    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                });
                
                const url = URL.createObjectURL(out);
                const a = document.createElement('a');
                a.href = url;
                const city = processedData.city || 'N/A';
                a.download = `دعوة-${city}-${workLocation}-${division}.docx`;
                document.body.appendChild(a);
                a.click();
                URL.revokeObjectURL(url);
                a.remove();
            }
            alert('تم إنشاء المستندات بنجاح!');
        } catch (error) {
            console.error('An unexpected error occurred during Word generation:', error);
            alert(`حدث خطأ غير متوقع أثناء إنشاء الملف: ${error.message}`);
        }
    });

    // --- Initial Load ---
    loadEmployeeData();
    loadInvitedList();
});