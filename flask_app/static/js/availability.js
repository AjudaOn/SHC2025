document.addEventListener('DOMContentLoaded', function () {
    const mhexField =
        document.getElementById('inputMhex') ||
        document.getElementById('mhex') ||
        document.querySelector('[name="mhex"]');
    const entradaField = document.getElementById('entrada');
    const saidaField = document.getElementById('saida');
    const uhField =
        document.getElementById('inputUH') ||
        document.getElementById('uh') ||
        document.querySelector('[name="uh"]');

    // Get reservation ID if we are editing (from URL)
    // Matches:
    // - /reservas/editar/<id>/
    // - /reservas/editar/admin/<id>/
    const match = window.location.pathname.match(/\/reservas\/editar\/(?:admin\/)?(\d+)/);
    let excludeId = match ? match[1] : null;

    console.log('Availability Script Loaded');
    console.log('Exclude ID:', excludeId);

    const initialMhex = mhexField ? (mhexField.value || '') : '';

    function updateAvailableUHs() {
        const mhex = (mhexField && mhexField.value) ? mhexField.value : '';
        const entrada = entradaField.value;
        const saida = saidaField.value;

        console.log('Checking availability for:', { mhex, entrada, saida });

        if (mhex && entrada && saida) {
            let url = `/api/available-uhs/?mhex=${encodeURIComponent(mhex)}&entrada=${encodeURIComponent(entrada)}&saida=${encodeURIComponent(saida)}`;
            if (excludeId) {
                url += `&exclude_id=${encodeURIComponent(excludeId)}`;
            }

            fetch(url)
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`HTTP ${response.status}`);
                    }
                    return response.json();
                })
                .then(data => {
                    // Clear current options
                    const currentVal = uhField.value;
                    uhField.innerHTML = '';

                    // Add default option
                    const defaultOption = document.createElement('option');
                    defaultOption.text = '---------';
                    defaultOption.value = '';
                    uhField.add(defaultOption);

                    data.available_uhs.forEach(uh => {
                        const option = document.createElement('option');
                        option.value = uh[0];
                        option.text = uh[1];
                        uhField.add(option);
                    });

                    // Restore selected value if it's still available;
                    // if not, keep it as an explicit option (common when editing).
                    if (currentVal) {
                        const values = Array.from(uhField.options).map(o => o.value);
                        if (!values.includes(currentVal) && mhex === initialMhex) {
                            const option = document.createElement('option');
                            option.value = currentVal;
                            option.text = currentVal;
                            uhField.add(option);
                        }
                        if (values.includes(currentVal) || mhex === initialMhex) {
                            uhField.value = currentVal;
                        } else {
                            uhField.value = '';
                        }
                    }
                })
                .catch(error => console.error('Error fetching UHs:', error));
        }
    }

    if (mhexField && entradaField && saidaField && uhField) {
        mhexField.addEventListener('change', updateAvailableUHs);
        entradaField.addEventListener('change', updateAvailableUHs);
        saidaField.addEventListener('change', updateAvailableUHs);

        // Initial check if fields are populated (Edit mode)
        if (mhexField.value && entradaField.value && saidaField.value) {
            // Optional: Trigger update on load, but be careful not to clear existing selection 
            // before we know if it's valid. The logic above handles restoring 'currentVal'.
            updateAvailableUHs();
        }
    }
});
