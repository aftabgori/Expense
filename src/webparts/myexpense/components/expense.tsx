import * as React from 'react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { TextField } from 'office-ui-fabric-react';
import { Button } from '@mui/material';
import DeleteIcon from '@mui/icons-material/Delete';
import AddCircleIcon from '@mui/icons-material/AddCircle';
import axios from 'axios';
// import { sp } from '@pnp/sp';
import { sp } from "@pnp/sp/presets/all";

interface IResource {
    Id: number;
    ResourceName1: string;
    Salary: number;
}

// interface Isections {
//     resourceName: string, timePeriod: string, allocation: string, fetchedSalary?: number
// }

interface IDropdownsProps {
    onResourceSelectionChange: (selectedResources: string[]) => void;
    onTimePeriodChange: (selectedTimePeriod: string) => void;
    onAllocationChange: (selectedAllocation: string) => void;
}

const Dropdowns: React.FC<IDropdownsProps> = ({

}) => {
    const [allocatedResources, setAllocatedResources] = React.useState<IDropdownOption[]>([]);
    const [totalExpense, setTotalExpense] = React.useState<number>(0);
    // const [resourceData, setResourceData] = React.useState<IResource[]>([]);
    const [ProjectName, setProjectName] = React.useState<string>('');
    const [sections, setSections] = React.useState<Array<{ resourceName: string, timePeriod: string, allocation: string, fetchedSalary?: number }>>(
        [{ resourceName: '', timePeriod: '', allocation: '' }]
    );
    const [isLoading, setIsLoading] = React.useState(false);
    const [showTable, setShowTable] = React.useState(false);
    const [showDropdowns, setShowDropdowns] = React.useState(true);
    const [error, setError] = React.useState<string>('');
    const [successMessage, setSuccessMessage] = React.useState<string>('');
    // const [activeIndex, setActiveIndex] = React.useState(-1);

    React.useEffect(() => {
        fetchAllocatedResources();
    }, []);

    const sharepointListName = 'ResourceRequests';
    const siteUrl = 'https://tuliptechcom.sharepoint.com/sites/poc/';

    sp.setup({
        sp: {
            baseUrl: siteUrl,
        },
    });

    const fetchAllocatedResources = async () => {
        try {
            const response = await axios.get(
                `https://tuliptechcom.sharepoint.com/sites/poc/_api/web/lists/getbytitle('Resource')/items?$select=ResourceName1`
            );

            const data: IResource[] = response.data.value;
            const options: IDropdownOption[] = data.map((resource: IResource) => ({
                key: resource.ResourceName1,
                text: resource.ResourceName1,
            }));
            setAllocatedResources(options);
            //setResourceData(data);
        } catch (error) {
            console.error('Error fetching resources:', error);
        }
    };

    const handleProjectNameChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        setProjectName(event.target.value);
    };

    const handleResourceChange = (index: number, selectedResource: string) => {
        const updatedSections = [...sections];
        updatedSections[index].resourceName = selectedResource;
        setSections(updatedSections);
    };

    const handleTimePeriodChange = (index: number, selectedTimePeriod: string) => {
        const updatedSections = [...sections];
        updatedSections[index].timePeriod = selectedTimePeriod;
        setSections(updatedSections);
    };

    const handleAllocationChange = (index: number, selectedAllocation: string) => {
        const updatedSections = [...sections];
        updatedSections[index].allocation = selectedAllocation;
        setSections(updatedSections);
    };

    const calculateTotalExpense = async () => {
        setIsLoading(true);
        setShowTable(false);

        let total = 0;
        let hasError = false;

        for (const section of sections) {
            if (!section.resourceName || !section.timePeriod || !section.allocation) {
                hasError = true;
                setError("Please fill in all the fields before calculating total expense.");
                setTimeout(() => {
                    setError('');
                }, 2000);
                break;
            }
            try {
                const response = await axios.get(
                    `https://tuliptechcom.sharepoint.com/sites/poc/_api/web/lists/getbytitle('Salary')/items?$select=Salary&$filter=EmployeeName eq '${section.resourceName}'&$orderby=Modified desc&$top=1`
                );
                const salaryEntries: any[] = response.data.value;

                if (salaryEntries.length > 0) {
                    const latestSalary = salaryEntries[0].Salary;
                    section.fetchedSalary = latestSalary;
                    const allocationPercentage = parseFloat(section.allocation.replace('%', '')) / 100;
                    const timePeriodInWeeks = convertToWeeks(section.timePeriod);
                    const fractionOfMonth = timePeriodInWeeks / 4;
                    const expense = (latestSalary * allocationPercentage) * fractionOfMonth;
                    total += expense;
                } else {
                    console.log('No salary entry found for', section.resourceName);
                }

            } catch (error) {
                console.error('Error fetching latest salary:', error);
            }
        }

        if (hasError) {
            setIsLoading(false);
            setShowTable(false);
        } else {
            setTotalExpense(total);
            setIsLoading(false);
            setShowTable(true);
            setShowDropdowns(false);
            setError('');
        }
    };

    const convertToWeeks = (timePeriod: string): number => {
        const [value, unit] = timePeriod.split(' ');
        const parsedValue = parseFloat(value);

        if (isNaN(parsedValue)) {
            return 0;
        }

        switch (unit) {
            case 'day':
                return parsedValue / 7;
            case 'week':
                return parsedValue;
            case 'month':
                return parsedValue * 4;
            default:
                return 0;
        }
    };

    const timePeriodOptions: IDropdownOption[] = [
        { key: '1 day', text: '1 day' },
        { key: '1 week', text: '1 week' },
        { key: '1 month', text: '1 month' },
        { key: '2 month', text: '2 month' },
        { key: '3 month', text: '3 month' },
        { key: '4 month', text: '4 month' },
        { key: '5 month', text: '5 month' },
        { key: '6 month', text: '6 month' },
    ];

    const allocationOptions: IDropdownOption[] = [
        { key: '10%', text: '10%' },
        { key: '20%', text: '20%' },
        { key: '30%', text: '30%' },
        { key: '40%', text: '40%' },
        { key: '50%', text: '50%' },
        { key: '60%', text: '60%' },
        { key: '70%', text: '70%' },
        { key: '80%', text: '80%' },
        { key: '90%', text: '90%' },
        { key: '100%', text: '100%' },
    ];

    const sendRequestToSharePoint = async () => {
        try {
            const listItemData = {
                Title: 'New Request',
                // Resources: JSON.stringify(sections),
                Resources: JSON.stringify(sections),
                TotalExpense: totalExpense.toFixed(2),
                ProjectName: ProjectName
            };

            const allItems = await sp.web.lists.getByTitle(sharepointListName).items.select('Resources', 'TotalExpense').get();

            // checking similar request
            const existingRequest = allItems.find(item => {
                try {
                    const itemResources = JSON.parse(item.Resources);
                    return JSON.stringify(itemResources) === JSON.stringify(sections) && item.TotalExpense === listItemData.TotalExpense;
                } catch (error) {
                    return false;
                }
            });

            if (existingRequest) {
                setError('Request has already been submitted.')
                setTimeout(() => {
                    setError('');
                }, 2000)
            } else {

                const response = await sp.web.lists.getByTitle(sharepointListName).items.add(listItemData);

                console.log('Item created:', response);
                setSuccessMessage('Request has been sent successfully!')

                setTimeout(() => {
                    setSuccessMessage('');

                }, 2000)
            }
        } catch (error) {
            console.error('Error creating item:', error);
            setError('An Error occurred while sending the request.')
        }
    };

    const addSection = () => {
        setSections([...sections, { resourceName: '', timePeriod: '', allocation: '' }]);
        //setActiveIndex(sections.length);
    };

    const removeSection = (index: number) => {
        const updatedSections = sections.filter((_, i) => i !== index);
        setSections(updatedSections);
        // calculateTotalExpense();
    };

    const resetAll = () => {
        setSections([{ resourceName: '', timePeriod: '', allocation: '' }]);
        setTotalExpense(0);
        setShowTable(false);
        setShowDropdowns(true);
        setProjectName('');
    };

    return (
        <div>
            <div style={{}}>
                {showDropdowns && (
                    <div style={{ boxShadow: '0 4px 4px 0 rgba(0, 0, 0, 0.2), 0 25px 50px 0 rgba(0, 0, 0, 0.1)', padding: '15px', color: '#333333', justifyContent: 'center' }}>
                        {sections.map((section, index) => (
                            <div key={index} style={{ display: 'flex', gap: '10px', width: '42rem' }}>
                                {index === 0 && (
                                    <div style={{ width: '100%' }}>
                                        <TextField
                                            label="Project Name"
                                            value={ProjectName}
                                            onChange={handleProjectNameChange}
                                        />
                                    </div>
                                )}
                                <div style={{ width: '100%' }}>
                                    <Dropdown
                                        label="Select Resource"
                                        options={allocatedResources}
                                        onChange={(event, option) => handleResourceChange(index, option?.text || '')}
                                    />
                                </div>
                                <div style={{ width: '100%' }}>
                                    <Dropdown
                                        label="Select Time Period"
                                        selectedKey={section.timePeriod}
                                        options={timePeriodOptions}
                                        onChange={(event, option) => handleTimePeriodChange(index, option?.text || '')}
                                    />
                                </div>
                                <div style={{ width: '100%' }}>
                                    <Dropdown
                                        label="Select Allocation"
                                        selectedKey={section.allocation}
                                        options={allocationOptions}
                                        onChange={(event, option) => handleAllocationChange(index, option?.text || '')}
                                    />
                                </div>
                                {index === sections.length - 1 && (
                                    <AddCircleIcon style={{ alignSelf: 'center', fontSize: '35px', marginTop: '23px', cursor: 'pointer' }} onClick={addSection} />
                                )}
                                {index > 0 && <DeleteIcon style={{ marginTop: '29px', fontSize: '32px', cursor: 'pointer' }} onClick={() => removeSection(index)} />}
                            </div>
                        ))}
                        <div style={{ textAlign: 'center', marginTop: '10px' }}>
                            <Button variant="contained" onClick={calculateTotalExpense}>Calculate Total Expense</Button>
                        </div>
                    </div>
                )}
                {isLoading && <div style={{ textAlign: 'center' }}>Calculating...</div>}
                {showTable && (
                    <div style={{
                        boxShadow: '0 4px 4px 0 rgba(0, 0, 0, 0.2), 0 25px 50px 0 rgba(0, 0, 0, 0.1)', padding: '15px', color: '#333333', justifyContent: 'center'
                    }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid black' }}>
                            <thead>
                                <tr>
                                    <th style={{ border: '1px solid black', padding: '5px' }}>Resource</th>
                                    <th style={{ border: '1px solid black', padding: '5px' }}>Time Period</th>
                                    <th style={{ border: '1px solid black', padding: '5px' }}>Allocation</th>
                                    <th style={{ border: '1px solid black', padding: '5px' }}>Expense</th>
                                </tr>
                            </thead>
                            <tbody>
                                {sections.map((section, index) => {
                                    const allocationPercentage = parseFloat(section.allocation.replace('%', '')) / 100;
                                    const expense = section.fetchedSalary ?
                                        (section.fetchedSalary * allocationPercentage) * (convertToWeeks(section.timePeriod) / 4)
                                        : 0;

                                    console.log("Expense", expense.toFixed(2));
                                    console.log("Selected Resource Data", section.fetchedSalary);

                                    return (
                                        <tr key={index}>
                                            <td style={{ border: '1px solid black', padding: '5px' }}>{section.resourceName}</td>
                                            <td style={{ border: '1px solid black', padding: '5px' }}>{section.timePeriod}</td>
                                            <td style={{ border: '1px solid black', padding: '5px' }}>{section.allocation}</td>
                                            <td style={{ border: '1px solid black', padding: '5px' }}>{expense.toFixed(2)}</td>
                                        </tr>
                                    );
                                })}
                                <tr>
                                    <td colSpan={3} style={{ border: '2px solid black', padding: '5px', fontSize: "20px", fontWeight: 'bold' }}>Total Expense</td>
                                    <td style={{ border: '2px solid black', padding: '5px', fontSize: "20px", fontWeight: 'bold' }}>{totalExpense.toFixed(2)}</td>
                                </tr>
                            </tbody>
                        </table>
                        <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'center', justifyContent: 'center' }}>
                            <div style={{ margin: '20px', backgroundColor: '#f0f0f0', padding: '10px', fontSize: '24px', width: 'auto' }}>
                                <div>Total Expense: {totalExpense.toFixed(2)}</div>
                            </div>
                            <div><Button style={{ height: '50px', width: '100px' }} variant="contained" onClick={resetAll}>Reset</Button></div>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-evenly' }}>
                            <div><Button variant="contained" onClick={() => sendRequestToSharePoint()}>Send Request</Button></div>
                            <div><Button variant="contained" >Allocate Resources</Button></div>
                        </div>
                    </div>
                )}
            </div>
            <div>
                {error && (
                    <div style={{ color: 'red', textAlign: 'center', marginTop: '10px' }}>
                        {error}
                    </div>
                )}
            </div>
            <div>
                {successMessage && (
                    <div style={{ color: 'red', textAlign: 'center', marginTop: '10px' }}>
                        {successMessage}
                    </div>
                )}
            </div>
        </div >
    );
};

export default Dropdowns;
