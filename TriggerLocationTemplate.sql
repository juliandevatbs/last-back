-- Trigger that executes when insert a new template 
CREATE TRIGGER trg_increase_quantity_templates_location
ON Template
AFTER INSERT
AS
BEGIN
	--UPDATE COUNTER ON CLIENT TABLE (TEMPLATES QUANTITY)
	UPDATE Location
	SET template_quantity = template_quantity + 1
	FROM Client c
	INNER JOIN inserted i ON c.id = i.client_id
END;





-- Decrease the number of templates when a template is deleted
CREATE TRIGGER trg_decrease_quantity_templates_location
ON Template
AFTER DELETE
AS BEGIN
	UPDATE Location
	SET template_quantity = template_quantity - 1
	FROM Client c
	INNER JOIN deleted d ON c.id = d.client_id
	-- Make sure it´s not negative
	WHERE c.template_quantity > 0
END;