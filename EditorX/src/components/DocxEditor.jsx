import React, { useState, useEffect, useRef } from 'react';
import '../App.css';

const DocxEditor = ({ content, onSave, onCancel }) => {
    const [editableContent, setEditableContent] = useState(content);
    const [focusedElement, setFocusedElement] = useState(null);

    useEffect(() => {
        if (focusedElement) {
            const { element, position } = focusedElement;
            const range = document.createRange();
            const sel = window.getSelection();
            try {
                range.setStart(element.childNodes[0], Math.min(position, element.innerText.length));
                range.collapse(true);
                sel.removeAllRanges();
                sel.addRange(range);
            } catch (error) {
                console.error('Failed to set cursor position:', error.message);
            }
        }
    }, [editableContent, focusedElement]);

    const saveCursorPosition = (element) => {
        const sel = window.getSelection();
        if (sel.rangeCount) {
            const range = sel.getRangeAt(0);
            const position = range.startOffset;
            setFocusedElement({ element, position });
        }
    };

    const handleTextChange = (event, itemIndex, runIndex, type, rowIndex, cellIndex, paraIndex) => {
        const newContent = [...editableContent];
        
        if (type === 'paragraph') {
            newContent[itemIndex].runs[runIndex].text = event.target.innerText;
        } else if (type === 'table') {
            newContent[itemIndex].rows[rowIndex][cellIndex].paragraphs[paraIndex].runs[runIndex].text = event.target.innerText;
        }
        
        saveCursorPosition(event.target);
        setEditableContent(newContent);
    };

    const handleSaveClick = () => {
        onSave(editableContent);
    };

    const handleCancelClick = () => {
        onCancel();
    };

    const renderEditableText = (run, itemIndex, runIndex, type, rowIndex, cellIndex, paraIndex) => (
        <span
            key={runIndex}
            contentEditable
            suppressContentEditableWarning
            style={{
                fontWeight: run.bold ? 'bold' : 'normal',
                fontStyle: run.italic ? 'italic' : 'normal',
                textDecoration: run.underline ? 'underline' : 'none',
                fontFamily: run.font.name || 'inherit',
                fontSize: run.font.size ? `${run.font.size}pt` : 'inherit',
                outline: 'none'
            }}
            onInput={(event) => handleTextChange(event, itemIndex, runIndex, type, rowIndex, cellIndex, paraIndex)}
            onFocus={(event) => saveCursorPosition(event.target)}
        >
            {run.text}
        </span>
    );

    return (
        <div>
            {editableContent.map((item, itemIndex) => {
                if (item.type === 'paragraph') {
                    return (
                        <p key={itemIndex}>
                            {item.runs.map((run, runIndex) =>
                                renderEditableText(run, itemIndex, runIndex, 'paragraph')
                            )}
                        </p>
                    );
                } else if (item.type === 'table') {
                    return (
                        <table key={itemIndex}>
                            <tbody>
                                {item.rows.map((row, rowIndex) => (
                                    <tr key={rowIndex}>
                                        {row.map((cell, cellIndex) => (
                                            <td key={cellIndex}>
                                                {cell.paragraphs.map((para, paraIndex) => (
                                                    <div
                                                        key={paraIndex}
                                                        style={{ outline: 'none', direction: 'ltr' }}
                                                    >
                                                        {para.runs.map((run, runIndex) =>
                                                            renderEditableText(run, itemIndex, runIndex, 'table', rowIndex, cellIndex, paraIndex)
                                                        )}
                                                    </div>
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
            })}
            <div>
                <button onClick={handleSaveClick}>Save Changes</button>
                <button onClick={handleCancelClick}>Cancel</button>
            </div>
        </div>
    );
};

export default DocxEditor;
