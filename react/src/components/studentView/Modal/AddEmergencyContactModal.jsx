import React, { useState, useEffect } from 'react';
import { Phone } from 'lucide-react';
import Modal from '../../utility/Modal';
import { formatPhoneNumber } from '../../utility/Conversion';

// Modal for adding an emergency contact.
// Collects a name, phone number, and relationship.
// NOTE: No persistence yet — onAdd is called with the entered values so the
// saving logic can be wired up later.
const AddEmergencyContactModal = ({ isOpen, onClose, onAdd }) => {
  const [name, setName] = useState('');
  const [number, setNumber] = useState('');
  const [relationship, setRelationship] = useState('');

  // Reset the fields whenever the modal is (re)opened
  useEffect(() => {
    if (isOpen) {
      setName('');
      setNumber('');
      setRelationship('');
    }
  }, [isOpen]);

  const isValid = name.trim() && number.trim();

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!isValid) return;
    // No saving yet — hand the values off to the caller (to be wired up later)
    if (onAdd) {
      onAdd({
        name: name.trim(),
        number: number.trim(),
        relationship: relationship.trim()
      });
    }
    if (onClose) onClose();
  };

  const labelStyle = {
    display: 'block',
    fontSize: 13,
    fontWeight: 600,
    color: '#374151',
    marginBottom: 4
  };
  const inputStyle = {
    width: '100%',
    padding: '8px 10px',
    fontSize: 14,
    color: '#1f2937',
    background: '#fff',
    border: '1px solid #d1d5db',
    borderRadius: 8,
    outline: 'none',
    boxSizing: 'border-box'
  };

  return (
    <Modal
      isOpen={isOpen}
      onClose={onClose}
      padding="20px"
      borderRadius="14px"
      style={{ maxWidth: 380, width: '90vw', alignItems: 'stretch' }}
    >
      <div style={{ width: '100%' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 16 }}>
          <span
            style={{
              width: 32,
              height: 32,
              borderRadius: '9999px',
              background: '#fecaca',
              color: '#b91c1c',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              flexShrink: 0
            }}
          >
            <Phone size={18} />
          </span>
          <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#1f2937' }}>
            Add Emergency Contact
          </h3>
        </div>

        <form onSubmit={handleSubmit}>
          <div style={{ marginBottom: 12 }}>
            <label htmlFor="emergency-name" style={labelStyle}>Name</label>
            <input
              id="emergency-name"
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="Full name"
              style={inputStyle}
              autoFocus
            />
          </div>

          <div style={{ marginBottom: 12 }}>
            <label htmlFor="emergency-number" style={labelStyle}>Number</label>
            <input
              id="emergency-number"
              type="tel"
              value={number}
              onChange={(e) => setNumber(formatPhoneNumber(e.target.value))}
              placeholder="Phone number"
              style={inputStyle}
            />
          </div>

          <div style={{ marginBottom: 20 }}>
            <label htmlFor="emergency-relationship" style={labelStyle}>Relationship</label>
            <input
              id="emergency-relationship"
              type="text"
              value={relationship}
              onChange={(e) => setRelationship(e.target.value)}
              placeholder="e.g. Parent, Spouse, Sibling"
              style={inputStyle}
            />
          </div>

          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
            <button
              type="button"
              onClick={onClose}
              style={{
                padding: '8px 16px',
                fontSize: 14,
                fontWeight: 600,
                color: '#374151',
                background: '#f3f4f6',
                border: 'none',
                borderRadius: 8,
                cursor: 'pointer'
              }}
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={!isValid}
              style={{
                padding: '8px 16px',
                fontSize: 14,
                fontWeight: 600,
                color: '#fff',
                background: isValid ? '#ef4444' : '#fca5a5',
                border: 'none',
                borderRadius: 8,
                cursor: isValid ? 'pointer' : 'not-allowed'
              }}
            >
              Add
            </button>
          </div>
        </form>
      </div>
    </Modal>
  );
};

export default AddEmergencyContactModal;
