import React, { useEffect, useState } from 'react';
import axios from 'axios';
import '../App.css';
import DocxEditor from './DocxEditor';

const DocxViewer = ({ filename }) => {
    const [content, setContent] = useState([]);
    const [originalContent, setOriginalContent] = useState([]);
    const [editing, setEditing] = useState(false);
    const [error, setError] = useState(null);

    useEffect(() => {
        const fetchContent = async () => {
            try {
                const response = await axios.get(`http://127.0.0.1:5000/get_document/${filename}`);
                setContent(response.data.content);
                setOriginalContent(response.data.content);
                setError(null);
            } catch (error) {
                console.error('Error fetching document content:', error);
                setError('Error fetching document content. Please try again.');
            }
        };

        fetchContent();
    }, [filename]);

    const handleEditClick = () => {
        setEditing(true);
    };

    const handleSaveClick = async (newContent) => {
        try {
            const response = await axios.post(`http://127.0.0.1:5000/save_document/${filename}`, {
                content: newContent
            });

            if (response.data && response.data.content) {
                setContent(response.data.content);
                setOriginalContent(response.data.content);
                setError(null);
                setEditing(false);
                console.log(response.data.message);
            } else {
                setError('Invalid response format from server.');
            }
        } catch (error) {
            console.error('Error saving document content:', error);
            setError('Error saving document content. Please try again.');
        }
    };

    const handleCancelClick = () => {
        setEditing(false);
        setError(null);
    };

    return (
        <div>
            {error && <div className="error-message">{error}</div>}
            <div className="edit-button">
                {!editing && <button onClick={handleEditClick}>Edit</button>}
            </div>
            <div className="a4-container">
                {editing ? (
                    <DocxEditor
                        content={content}
                        onSave={handleSaveClick}
                        onCancel={handleCancelClick}
                    />
                ) : (
                    content.map((item, index) => {
                        if (item.type === 'paragraph') {
                            return (
                                <p key={index}>
                                    {item.runs.map((run, runIndex) => (
                                        <span
                                            key={runIndex}
                                            style={{
                                                fontWeight: run.bold ? 'bold' : 'normal',
                                                fontStyle: run.italic ? 'italic' : 'normal',
                                                textDecoration: run.underline ? 'underline' : 'none',
                                                fontFamily: run.font.name || 'inherit',
                                                fontSize: run.font.size ? `${run.font.size}pt` : 'inherit'
                                            }}
                                        >
                                            {run.text}
                                        </span>
                                    ))}
                                </p>
                            );
                        } else if (item.type === 'table') {
                            return (
                                <table key={index}>
                                    <tbody>
                                        {item.rows.map((row, rowIndex) => (
                                            <tr key={rowIndex}>
                                                {row.map((cell, cellIndex) => (
                                                    <td key={cellIndex}>
                                                        {cell.paragraphs.map((para, paraIndex) => (
                                                            <p key={paraIndex}>
                                                                {para.runs.map((run, runIndex) => (
                                                                    <span
                                                                        key={runIndex}
                                                                        style={{
                                                                            fontWeight: run.bold ? 'bold' : 'normal',
                                                                            fontStyle: run.italic ? 'italic' : 'normal',
                                                                            textDecoration: run.underline ? 'underline' : 'none',
                                                                            fontFamily: run.font.name || 'inherit',
                                                                            fontSize: run.font.size ? `${run.font.size}pt` : 'inherit'
                                                                        }}
                                                                    >
                                                                        {run.text}
                                                                    </span>
                                                                ))}
                                                            </p>
                                                        ))}
                                                    </td>
                                                ))}
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            );
                        }
                        return null;
                    })
                )}
            </div>
        </div>
    );
};

export default DocxViewer;
